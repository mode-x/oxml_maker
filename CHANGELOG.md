## [Unreleased]

## [0.1.1] - 2025-10-20

### Added
- Rubyzip 3.2 support for DOCX generation
- Comprehensive test suite (63 tests, 274 assertions)
- Rails-optional architecture with smart public directory detection  
- Full workflow integration tests
- ZIP functionality tests
- Complete documentation consolidation

### Changed
- Updated rubyzip dependency from ~> 2.3 to ~> 3.2
- Improved method naming (`copy_zip_to_docx` → `convert_zip_to_docx`)
- Extracted section generation logic into separate method
- Consolidated TESTING.md and EXAMPLE.md into README.md

### Fixed
- Missing ZIP-to-DOCX conversion step in Document#create
- Rubyzip 3.x API compatibility (removed deprecated CREATE constant)
- Rails dependency made optional with fallback support

## [0.1.0] - 2025-06-09

- Initial release
