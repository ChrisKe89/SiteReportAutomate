# Changelog

## [0.1.7] - 2025-10-22
### Added
- Documented the report download workflow and default output path for the cleaned
  firmware device list.
- Added a dedicated `fetch_ast_toner.py` automation that exports lookup results
  to `AST_OUTPUT_CSV` based on the product family mappings in `RDHC.html`.

### Changed
- Updated `schedule_firmware.py` to load devices from CSV as well as Excel,
  write structured JSON logs when requested, and require a valid
  `FIRMWARE_STORAGE_STATE` before launching the browser.
- Normalised environment variable handling across the scripts so Windows-style
  paths are safely resolved on any platform.

## [0.1.6] - 2025-10-21
### Fixed
- Ensure the firmware scheduler presses the Reset button after a successful submission so the form is ready for the next device.

## [0.1.5] - 2025-10-21
### Added
- Screenshot capture for each toner lookup step to provide a visual audit trail.

### Fixed
- Verified toner portal inputs before submitting searches so required fields are
  consistently populated.

## [0.1.4] - 2025-10-20
### Added
- Optional `FIRMWARE_HTTP_USERNAME` / `FIRMWARE_HTTP_PASSWORD` environment variables and a warm-up navigation hook to support portals that require HTTP authentication before loading the scheduling form.
- Guidance in the README and `.env.example` for configuring the new firmware scheduler options.

### Fixed
- Hardened `schedule_firmware.py` against missing active worksheets and broadened row parsing so mypy recognises the worksheet API.
- Display a helpful remediation hint when Playwright reports `ERR_INVALID_AUTH_CREDENTIALS` during navigation.

## [0.1.3] - 2024-05-23
### Fixed
- Allow `ast_toner.py` to proceed when `storage_state.json` is missing by making the storage state optional and configurable via `AST_TONER_STORAGE_STATE`.
- Added documentation for the new storage state resolution workflow.

## [0.1.2] - 2024-05-22
### Fixed
- Automatically locate the latest cleaned workbook in `downloads/` to prevent missing-file crashes in `ast_toner.py`.
- Tightened workbook typing so Pylance recognises worksheet methods.

## [0.1.1] - 2024-05-21
### Added
- Hardened `ast_toner.py` with structured logging, resilient parsing, and safeguards for missing workbook columns.
- Documented toner status automation workflow in the README and set explicit release metadata.

### Fixed
- Corrected `requirements.txt` to pin BeautifulSoup with an explicit version and a terminating newline.

## [0.1.0] - 2024-05-20
### Added
- Initial automation scripts for EP Gateway business rules and toner status lookups.
