import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from process.extract_data import PortalPms as PortalPmsExtract, STORAGE_FULL
from process.transform_data import CloudPmsTrans, PortalPmsTrans
from process.refresh_excel_and_send import RefreshAndSend
import config
from services.logger import get_logger
from services.email import Email

# Replaced in run() so each system writes its own main_<system> log file.
logger = logging.getLogger(__name__)

MAX_RETRY = 2
SYSTEMS = ('portal_pms', 'cloud_pms')

# ---------------------------------------------------------------------------
# Per-site runners
# ---------------------------------------------------------------------------

def _step_failed(status) -> bool:
    return bool(status and 'FAIL' in status.upper())


def _storage_full(status) -> bool:
    return bool(status and STORAGE_FULL in status)


def _cloud_pms_transform_and_send(country: str, site_id: str) -> dict:
    """Run transform → refresh&send for one CloudPms site. Stops at the first failed step."""
    result = {'site_id': site_id}
    logger.info(f"[{country}/{site_id}] transform: running...")
    print(f"[{country}/{site_id}] transform: running...")
    _, transform_status = CloudPmsTrans().run(country, site_id)
    result['transform'] = transform_status
    logger.info(f"[{country}/{site_id}] transform: {transform_status}")
    print(f"[{country}/{site_id}] transform: {transform_status}")

    if _step_failed(transform_status):
        result['refresh_send'] = 'REFRESH AND SEND EMAIL SKIPPED - transform did not succeed'
        logger.info(f"[{country}/{site_id}] refresh_send: skipped - transform did not succeed")
        print(f"[{country}/{site_id}] refresh_send: skipped - transform did not succeed")
        return result

    logger.info(f"[{country}/{site_id}] refresh_send: running...")
    print(f"[{country}/{site_id}] refresh_send: running...")
    _, send_status = RefreshAndSend().run(country, site_id)
    result['refresh_send'] = send_status
    logger.info(f"[{country}/{site_id}] refresh_send: {send_status}")
    print(f"[{country}/{site_id}] refresh_send: {send_status}")
    return result


def _portal_pms_transform_and_send(country: str, site_id: str) -> tuple[str, str]:
    """Run transform → refresh&send for one configured PortalPms site."""
    logger.info(f"[{country}/{site_id}] transform: running...")
    print(f"[{country}/{site_id}] transform: running...")
    _, transform_status = PortalPmsTrans().run(country, site_id)
    logger.info(f"[{country}/{site_id}] transform: {transform_status}")
    print(f"[{country}/{site_id}] transform: {transform_status}")

    if _step_failed(transform_status):
        send_status = 'REFRESH AND SEND EMAIL SKIPPED - transform did not succeed'
        logger.info(f"[{country}/{site_id}] refresh_send: skipped - transform did not succeed")
        print(f"[{country}/{site_id}] refresh_send: skipped - transform did not succeed")
        return transform_status, send_status

    logger.info(f"[{country}/{site_id}] refresh_send: running...")
    print(f"[{country}/{site_id}] refresh_send: running...")
    _, send_status = RefreshAndSend().run(country, site_id)
    logger.info(f"[{country}/{site_id}] refresh_send: {send_status}")
    print(f"[{country}/{site_id}] refresh_send: {send_status}")
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

def _process_cloud_pms(max_workers: int = 3) -> list[dict]:
    """All CloudPms sites: transform + refresh&send fully concurrent (no extract needed)."""
    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {
            executor.submit(_cloud_pms_transform_and_send, country, site_id): (country, site_id)
            for country, site_id in config.group_ids['cloud_pms']
        }
        for future in as_completed(future_map):
            _, site_id = future_map[future]
            try:
                result = future.result()
            except Exception as e:
                logger.error(f"{site_id}: Unexpected error - {e}")
                print(f"[{site_id}] Unexpected error - {e}")
                result = {
                    'site_id': site_id,
                    'transform': f'TRANSFORM DATA FAILED - {e}',
                    'refresh_send': 'REFRESH AND SEND EMAIL FAILED - upstream error',
                }
            results.append(result)
    return results


def _process_portal_pms(max_workers: int = 3) -> list[dict]:
    """
    PortalPms pipeline:
      - extract runs sequentially (shared Chrome profile)
      - if extract fails, transform+send are skipped and the site moves on immediately
      - otherwise transform+send is submitted to the thread pool right after extract
      - multiple transform+send jobs run concurrently across sites
    """
    results_map: dict[str, dict] = {}
    future_map: dict = {}

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for country, site_id in config.group_ids['portal_pms']:
            # --- Sequential extract (shared Chrome profile) ---
            logger.info(f"[{country}/{site_id}] extract: running...")
            print(f"[{country}/{site_id}] extract: running...")
            try:
                _, extract_status = PortalPmsExtract().run(country, site_id)
            except Exception as e:
                logger.error(f"{site_id}: Extract raised unexpected error - {e}")
                extract_status = f"EXTRACT DATA FAILED - {e}"
            logger.info(f"[{country}/{site_id}] extract: {extract_status}")
            print(f"[{country}/{site_id}] extract: {extract_status}")

            if _step_failed(extract_status) or _storage_full(extract_status):
                logger.info(f"[{country}/{site_id}] transform/refresh_send: skipped - extract did not succeed")
                print(f"[{country}/{site_id}] transform/refresh_send: skipped - extract did not succeed")
                results_map[site_id] = {
                    'site_id': site_id,
                    'extract': extract_status,
                    'transform': 'TRANSFORM DATA SKIPPED - extract did not succeed',
                    'refresh_send': 'REFRESH AND SEND EMAIL SKIPPED - extract did not succeed',
                }
                continue

            results_map[site_id] = {
                'site_id': site_id,
                'extract': extract_status,
                'transform': None,
                'refresh_send': None,
            }

            # --- Submit transform+send to thread pool immediately after extract ---
            future = executor.submit(_portal_pms_transform_and_send, country, site_id)
            future_map[future] = site_id

        # Collect transform+send results as they complete
        for future in as_completed(future_map):
            site_id = future_map[future]
            try:
                transform_status, send_status = future.result()
            except Exception as e:
                logger.error(f"{site_id}: Transform/Send raised unexpected error - {e}")
                print(f"[{site_id}] Transform/Send raised unexpected error - {e}")
                transform_status = f'TRANSFORM DATA FAILED - {e}'
                send_status = 'REFRESH AND SEND EMAIL FAILED - upstream error'
            results_map[site_id]['transform'] = transform_status
            results_map[site_id]['refresh_send'] = send_status

    return list(results_map.values())


def _retry_failed(failed: list[dict], portal_pms_map: dict, cloud_pms_map: dict, max_workers: int = 3) -> dict:
    """
    Retry a batch of failed sites, re-running only from each site's first failed step.

    PortalPms sites whose extract (or transform, which restarts from extract) failed
    must redo the extract sequentially — like the initial run, all PortalPms extracts
    share one Chrome profile, so concurrent extracts would conflict. Each such site's
    transform+send retry is submitted to the thread pool right after its extract
    succeeds, so it overlaps with the next site's extract. PortalPms sites failing only
    at send, and all CloudPms retries, need no extract and run fully concurrently.
    """
    results: dict[str, dict] = {}
    future_map: dict = {}

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for prev_result in failed:
            site_id = prev_result['site_id']
            is_portal_pms = site_id in portal_pms_map
            country = portal_pms_map.get(site_id) or cloud_pms_map.get(site_id)
            result = dict(prev_result)

            if not is_portal_pms:  # CloudPms — no extract step, always safe to run concurrently
                logger.info(f"[{country}/{site_id}] retrying from transform")
                print(f"[{country}/{site_id}] retrying from transform")
                future = executor.submit(_cloud_pms_transform_and_send, country, site_id)
                future_map[future] = (site_id, result, 'dict')
                continue

            extract_failed   = _step_failed(prev_result.get('extract'))
            transform_failed = _step_failed(prev_result.get('transform'))

            if extract_failed or transform_failed:
                reason = "retrying from extract" if extract_failed else "transform failed — retrying from extract"
                logger.info(f"[{country}/{site_id}] {reason}")
                print(f"[{country}/{site_id}] {reason}")
                try:
                    _, extract_status = PortalPmsExtract().run(country, site_id)
                except Exception as e:
                    extract_status = f"EXTRACT DATA FAILED - {e}"
                result['extract'] = extract_status
                logger.info(f"[{country}/{site_id}] extract: {extract_status}")
                print(f"[{country}/{site_id}] extract: {extract_status}")

                if _step_failed(extract_status):
                    results[site_id] = result
                    continue
                if _storage_full(extract_status):
                    # Clear the old FAIL statuses so this site is not retried again
                    result['transform'] = 'TRANSFORM DATA SKIPPED - extract did not succeed'
                    result['refresh_send'] = 'REFRESH AND SEND EMAIL SKIPPED - extract did not succeed'
                    results[site_id] = result
                    continue

                future = executor.submit(_portal_pms_transform_and_send, country, site_id)
                future_map[future] = (site_id, result, 'tuple')
                continue

            # Only send failed — no extract needed, safe to run concurrently
            logger.info(f"[{country}/{site_id}] refresh & send failed — retrying from transform")
            print(f"[{country}/{site_id}] refresh & send failed — retrying from transform")
            future = executor.submit(_portal_pms_transform_and_send, country, site_id)
            future_map[future] = (site_id, result, 'tuple')

        for future in as_completed(future_map):
            site_id, result, kind = future_map[future]
            try:
                outcome = future.result()
            except Exception as e:
                logger.error(f"{site_id}: Retry raised unexpected error - {e}")
                print(f"[{site_id}] Retry raised unexpected error - {e}")
                result['transform'] = f'TRANSFORM DATA FAILED - {e}'
                result['refresh_send'] = 'REFRESH AND SEND EMAIL FAILED - upstream error'
                results[site_id] = result
                continue

            if kind == 'dict':
                result.update(outcome)
            else:
                result['transform'], result['refresh_send'] = outcome
            results[site_id] = result

    return results


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


def _send_summary(results: list[dict], system: str) -> None:
    """Log and email an HTML summary table."""
    title = f"Daily Operation Report Run Summary - {system.capitalize()}"
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
        <h2>{title}</h2>
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
            'subject': title,
            'body': {'contentType': 'HTML', 'content': body},
            'toRecipients': [{'emailAddress': {'address': sender_email}}],
        },
        'saveToSentItems': 'true',
    }
    Email().send(sender_email, message)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def run(system: str) -> None:
    """Run, retry and summarise every site of one PMS system ('portal_pms' or 'cloud_pms')."""
    if system not in SYSTEMS:
        raise ValueError(f"Unknown system {system!r} - expected one of {SYSTEMS}")

    global logger
    logger = get_logger(f'main_{system}', config.paths['log_dir'])

    all_results: dict[str, dict] = {}

    # Each run only knows its own sites, so a retry never touches the other system.
    sites = {site_id: country for country, site_id in config.group_ids[system]}
    portal_pms_map = sites if system == 'portal_pms' else {}
    cloud_pms_map = sites if system == 'cloud_pms' else {}

    # --- Initial run ---
    process = _process_portal_pms if system == 'portal_pms' else _process_cloud_pms
    for result in process():
        all_results[result['site_id']] = result

    # --- Retry failed sites (concurrently, after all sites done) ---
    for attempt in range(1, MAX_RETRY + 1):
        failed = [r for r in all_results.values() if _is_failed(r)]
        if not failed:
            break

        failed_ids = [r['site_id'] for r in failed]
        logger.info(f"Retry attempt {attempt}/{MAX_RETRY} for: {failed_ids}")
        print(f"\nRetrying {len(failed)} failed site(s): {failed_ids}")

        all_results.update(_retry_failed(failed, portal_pms_map, cloud_pms_map))

    # Preserve config order in summary output
    ordered = [all_results[site_id] for site_id in sites if site_id in all_results]

    _send_summary(ordered, system)
