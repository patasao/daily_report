from concurrent.futures import ThreadPoolExecutor, as_completed
from process.extract_data import Guestline as GuestlineExtract
from process.transform_data import MewsTrans, GuestlineTrans
from process.refresh_excel_and_send import RefreshAndSend
import config
from services.logger import get_logger
from services.email import Email

logger = get_logger('main', config.paths['log_dir'])

MAX_RETRY = 2

# ---------------------------------------------------------------------------
# Per-site runners
# ---------------------------------------------------------------------------

def _mews_transform_and_send(country: str, site_id: str) -> dict:
    """Run transform → refresh&send for one Mews site. Returns result dict."""
    result = {'site_id': site_id}
    _, transform_status = MewsTrans().run(country, site_id)
    result['transform'] = transform_status
    _, send_status = RefreshAndSend().run(country, site_id)
    result['refresh_send'] = send_status
    return result


def _guestline_transform_and_send(country: str, site_id: str) -> tuple[str, str]:
    """Run transform → refresh&send for one Guestline site (called after extract)."""
    _, transform_status = GuestlineTrans().run(country, site_id)
    _, send_status = RefreshAndSend().run(country, site_id)
    return transform_status, send_status


def _is_failed(result: dict) -> bool:
    return any(
        v and 'FAIL' in v.upper()
        for k, v in result.items()
        if k != 'site_id' and v is not None
    )


# ---------------------------------------------------------------------------
# Batch processors
# ---------------------------------------------------------------------------

def _process_mews(max_workers: int = 3) -> list[dict]:
    """All Mews sites: transform + refresh&send fully concurrent (no extract needed)."""
    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {
            executor.submit(_mews_transform_and_send, country, site_id): (country, site_id)
            for country, site_id in config.group_ids['mews']
        }
        for future in as_completed(future_map):
            _, site_id = future_map[future]
            try:
                result = future.result()
            except Exception as e:
                logger.error(f"{site_id}: Unexpected error - {e}")
                result = {
                    'site_id': site_id,
                    'transform': f'TRANSFORM DATA FAILED - {e}',
                    'refresh_send': 'REFRESH AND SEND EMAIL FAILED - upstream error',
                }
            results.append(result)
    return results


def _process_guestline(max_workers: int = 3) -> list[dict]:
    """
    Guestline pipeline:
      - extract runs sequentially (shared Chrome profile)
      - immediately after each extract completes, submit transform+send to thread pool
      - multiple transform+send jobs run concurrently across sites
    """
    results_map: dict[str, dict] = {}
    future_map: dict = {}

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for country, site_id in config.group_ids['guestline']:
            # --- Sequential extract (shared Chrome profile) ---
            try:
                _, extract_status = GuestlineExtract().run(country, site_id)
            except Exception as e:
                logger.error(f"{site_id}: Extract raised unexpected error - {e}")
                extract_status = f"EXTRACT DATA FAILED - {e}"

            results_map[site_id] = {
                'site_id': site_id,
                'extract': extract_status,
                'transform': None,
                'refresh_send': None,
            }

            # --- Submit transform+send to thread pool immediately after extract ---
            future = executor.submit(_guestline_transform_and_send, country, site_id)
            future_map[future] = site_id

        # Collect transform+send results as they complete
        for future in as_completed(future_map):
            site_id = future_map[future]
            try:
                transform_status, send_status = future.result()
            except Exception as e:
                logger.error(f"{site_id}: Transform/Send raised unexpected error - {e}")
                transform_status = f'TRANSFORM DATA FAILED - {e}'
                send_status = 'REFRESH AND SEND EMAIL FAILED - upstream error'
            results_map[site_id]['transform'] = transform_status
            results_map[site_id]['refresh_send'] = send_status

    return list(results_map.values())


def _step_failed(status) -> bool:
    return bool(status and 'FAIL' in status.upper())


def _retry_site(country: str, site_id: str, is_guestline: bool, prev_result: dict) -> dict:
    """Re-run only from the first failed step for one site."""
    result = dict(prev_result)

    if is_guestline:
        extract_failed   = _step_failed(prev_result.get('extract'))
        transform_failed = _step_failed(prev_result.get('transform'))
        send_failed      = _step_failed(prev_result.get('refresh_send'))

        if extract_failed:
            # No step before extract — retry extract itself
            logger.info(f"{site_id}: retrying from extract")
            try:
                _, extract_status = GuestlineExtract().run(country, site_id)
            except Exception as e:
                extract_status = f"EXTRACT DATA FAILED - {e}"
            result['extract'] = extract_status
            if not _step_failed(extract_status):
                transform_status, send_status = _guestline_transform_and_send(country, site_id)
                result['transform'] = transform_status
                result['refresh_send'] = send_status

        elif transform_failed:
            # Retry from extract (step before transform)
            logger.info(f"{site_id}: transform failed — retrying from extract")
            try:
                _, extract_status = GuestlineExtract().run(country, site_id)
            except Exception as e:
                extract_status = f"EXTRACT DATA FAILED - {e}"
            result['extract'] = extract_status
            if not _step_failed(extract_status):
                transform_status, send_status = _guestline_transform_and_send(country, site_id)
                result['transform'] = transform_status
                result['refresh_send'] = send_status

        elif send_failed:
            # Retry from transform (step before send)
            logger.info(f"{site_id}: refresh & send failed — retrying from transform")
            transform_status, send_status = _guestline_transform_and_send(country, site_id)
            result['transform'] = transform_status
            result['refresh_send'] = send_status

    else:  # Mews
        transform_failed = _step_failed(prev_result.get('transform'))
        send_failed      = _step_failed(prev_result.get('refresh_send'))

        if transform_failed:
            # No step before transform for Mews — retry transform (and send)
            logger.info(f"{site_id}: retrying from transform")
            result = _mews_transform_and_send(country, site_id)
        elif send_failed:
            # Retry from transform (step before send)
            logger.info(f"{site_id}: refresh & send failed — retrying from transform")
            result = _mews_transform_and_send(country, site_id)

    return result


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

def _cell_color(status) -> str:
    if not status:
        return '#ffffff'
    if 'FAIL' in status.upper():
        return '#f8d7da'
    if 'SUCCESS' in status.upper():
        return '#d4edda'
    return '#fff3cd'


def _send_summary(results: list[dict]) -> None:
    """Log and email an HTML summary table."""
    print("\n=== SUMMARY REPORT ===")
    logger.info("=== SUMMARY REPORT ===")

    has_extract = any('extract' in r for r in results)

    headers = ['Site ID']
    if has_extract:
        headers.append('Extract')
    headers += ['Transform', 'Refresh & Send']

    header_html = ''.join(
        f"<th style='padding:12px;text-align:left;'>{h}</th>" for h in headers
    )

    rows_html = ''
    for r in results:
        stat_keys = (['extract'] if has_extract else []) + ['transform', 'refresh_send']
        cells = f"<td style='padding:8px;border:1px solid #ddd;'>{r['site_id']}</td>"
        cells += ''.join(
            f"<td style='padding:8px;border:1px solid #ddd;background:{_cell_color(r.get(k))};'>{r.get(k) or '-'}</td>"
            for k in stat_keys
        )
        rows_html += f"<tr>{cells}</tr>"

        line = f"{r['site_id']}: " + ' | '.join(f"{k}={r.get(k, '-')}" for k in stat_keys)
        print(line)
        logger.info(line)

    body = f"""
    <html><body style="font-family:Arial,sans-serif;">
        <h2>Daily Operation Report Run Summary</h2>
        <table style="border-collapse:collapse;width:100%;max-width:1100px;">
            <tr style="background:#4CAF50;color:white;">{header_html}</tr>
            {rows_html}
        </table>
    </body></html>
    """

    sender_email = config.e_mail['sender_email']
    if not sender_email:
        logger.error("Environment variable 'e_mail' is not set — summary email skipped.")
        return

    message = {
        'message': {
            'subject': 'Daily Operation Report Run Summary',
            'body': {'contentType': 'HTML', 'content': body},
            'toRecipients': [{'emailAddress': {'address': sender_email}}],
        },
        'saveToSentItems': 'true',
    }
    Email().send(sender_email, message)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    all_results: dict[str, dict] = {}

    mews_map = {site_id: country for country, site_id in config.group_ids['mews']}
    guestline_map = {site_id: country for country, site_id in config.group_ids['guestline']}

    # --- Initial run ---
    for result in _process_guestline():
        all_results[result['site_id']] = result

    for result in _process_mews():
        all_results[result['site_id']] = result

    # --- Retry failed sites (sequentially, after all sites done) ---
    for attempt in range(1, MAX_RETRY + 1):
        failed = [r for r in all_results.values() if _is_failed(r)]
        if not failed:
            break

        failed_ids = [r['site_id'] for r in failed]
        logger.info(f"Retry attempt {attempt}/{MAX_RETRY} for: {failed_ids}")
        print(f"\nRetrying {len(failed)} failed site(s): {failed_ids}")

        for r in failed:
            site_id = r['site_id']
            is_guestline = site_id in guestline_map
            country = guestline_map.get(site_id) or mews_map.get(site_id)
            all_results[site_id] = _retry_site(country, site_id, is_guestline, r)

    # Preserve config order in summary output
    ordered = []
    for _, site_id in config.group_ids['guestline']:
        if site_id in all_results:
            ordered.append(all_results[site_id])
    for _, site_id in config.group_ids['mews']:
        if site_id in all_results:
            ordered.append(all_results[site_id])

    _send_summary(ordered)


if __name__ == '__main__':
    main()
