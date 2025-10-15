# Changelog

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
