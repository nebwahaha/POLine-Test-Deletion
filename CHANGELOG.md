# Changelog

All notable changes to the Excel Cleaner project will be documented in this file.

## [1.0.0] - 2024

### Initial Release

#### Added
- ✅ Core Excel cleaning functionality
- ✅ Drag-and-drop support for Windows
- ✅ File picker dialog for browsing
- ✅ Pattern matching for 4 columns (H, I, BO, BV)
- ✅ Case-insensitive substring matching
- ✅ Safe operation (never overwrites original files)
- ✅ Output naming convention: `<filename>_CLEANED.xlsx`
- ✅ Success message with statistics
- ✅ Comprehensive error handling
- ✅ PyInstaller build configuration
- ✅ Automated build script (build.bat)
- ✅ Test script with sample data generation
- ✅ Complete documentation suite

#### Cleaning Rules Implemented
- **Column BV (ShipmentID)**: Remove rows containing "FOC"
- **Column H (Order)**: Remove rows containing "test", "testing", "M88", "GB Test", "GB Testing", "GB"
- **Column I (Buyer PO Number)**: Remove rows containing "test", "testing", "FOC"
- **Column BO (Comment)**: Remove rows containing "FOC", "M88"

#### Documentation
- README.md - Comprehensive project documentation
- BUILD_INSTRUCTIONS.txt - Quick build reference
- USER_GUIDE.txt - End-user manual
- PROJECT_SUMMARY.md - Technical overview
- START_HERE.txt - Quick start guide
- WORKFLOW_DIAGRAM.txt - Visual workflow representation
- CHANGELOG.md - This file

#### Technical Details
- Python 3.8+ compatible
- Uses pandas for efficient data processing
- Uses openpyxl for Excel file I/O
- Uses tkinter for GUI dialogs
- Single-file executable output
- No console window (windowed mode)
- Handles large files efficiently

---

## Future Enhancements (Planned)

### [1.1.0] - Potential Features
- [ ] Configuration file for custom rules
- [ ] Batch processing (multiple files at once)
- [ ] Preview mode (show what would be removed)
- [ ] Detailed cleaning report export
- [ ] Support for .xls (older Excel format)
- [ ] Custom output directory selection
- [ ] Logging functionality

### [1.2.0] - Advanced Features
- [ ] GUI for rule management
- [ ] Regex pattern support
- [ ] Column mapping customization
- [ ] Undo functionality
- [ ] Excel formula preservation
- [ ] Conditional formatting preservation
- [ ] Multi-sheet support

### [2.0.0] - Major Enhancements
- [ ] Web-based interface option
- [ ] Cloud storage integration
- [ ] Scheduled/automated cleaning
- [ ] API for programmatic access
- [ ] Database export options
- [ ] Advanced analytics and reporting

---

## Version History Template

### [Version Number] - YYYY-MM-DD

#### Added
- New features that have been added

#### Changed
- Changes to existing functionality

#### Deprecated
- Features that will be removed in future versions

#### Removed
- Features that have been removed

#### Fixed
- Bug fixes

#### Security
- Security-related changes

---

## Notes

- This project follows [Semantic Versioning](https://semver.org/)
- Version format: MAJOR.MINOR.PATCH
  - MAJOR: Incompatible API changes
  - MINOR: Backwards-compatible functionality additions
  - PATCH: Backwards-compatible bug fixes

---

## Contribution Guidelines

When updating this changelog:
1. Add new entries at the top (most recent first)
2. Use ISO date format (YYYY-MM-DD)
3. Group changes by type (Added, Changed, Fixed, etc.)
4. Be clear and concise
5. Link to issues/PRs when applicable
6. Keep the format consistent

---

**Current Version: 1.0.0**
**Status: Stable**
**Last Updated: 2024**
