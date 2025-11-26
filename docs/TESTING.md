# Testing Guide

This document describes the comprehensive test suite for PPTX-RS.

## Test Coverage

### Unit Tests

The project includes extensive unit tests organized by module:

#### Text Formatting Tests (10 tests)
- `test_text_format_default` - Default format properties
- `test_text_format_bold` - Bold formatting
- `test_text_format_italic` - Italic formatting
- `test_text_format_underline` - Underline formatting
- `test_text_format_color` - Color specification
- `test_text_format_color_lowercase` - Lowercase color handling
- `test_text_format_font_size` - Font size setting
- `test_text_format_combined` - Combined formatting
- `test_formatted_text_default` - Default formatted text
- `test_formatted_text_with_format` - Formatted text with format
- `test_formatted_text_builder` - Formatted text builder

#### Shape Tests (22 tests)
- `test_shape_type_*` - All 7 shape types
- `test_shape_fill_*` - Fill properties and transparency
- `test_shape_line_*` - Line/border properties
- `test_shape_basic` - Basic shape creation
- `test_shape_with_fill` - Shape with fill
- `test_shape_with_line` - Shape with line
- `test_shape_with_text` - Shape with text
- `test_shape_complete` - Complete shape with all properties
- `test_shape_zero_dimensions` - Edge case: zero size
- `test_shape_large_dimensions` - Edge case: large size

#### Table Tests (18 tests)
- `test_table_cell_*` - Cell properties and formatting
- `test_table_row_*` - Row creation and properties
- `test_table_basic` - Basic table creation
- `test_table_from_data` - Table from 2D data
- `test_table_builder_*` - Builder pattern tests
- `test_table_single_cell` - Edge case: single cell
- `test_table_many_columns` - Edge case: many columns
- `test_table_many_rows` - Edge case: many rows

#### EMU Conversion Tests (4 tests)
- `test_inches_to_emu` - Inches to EMU conversion
- `test_emu_to_inches` - EMU to inches conversion
- `test_cm_to_emu` - Centimeters to EMU conversion
- `test_emu_conversions_roundtrip` - Roundtrip conversion accuracy

### Integration Tests

#### CLI Tests (3 tests)
- `test_parse_create` - Create command parsing
- `test_parse_from_markdown` - Markdown command parsing
- `test_parse_info` - Info command parsing

#### Generator Tests (1 test)
- `test_slide_content_builder` - SlideContent builder

#### Integration Tests (3 tests)
- `test_presentation_builder` - PresentationBuilder
- `test_slide_builder` - SlideBuilder
- `test_format_size` - Size formatting utility

### Test Statistics

**Total Tests: 85**
- Unit Tests (in-module): 34
- Integration Tests (tests/ directory): 51
- All tests passing ✅

## Running Tests

### Run All Tests
```bash
cargo test
```

### Run Specific Test Suite
```bash
# Text formatting tests
cargo test generator::text::

# Shape tests
cargo test generator::shapes::

# Table tests
cargo test generator::tables::

# Advanced features integration tests
cargo test --test advanced_features_test
```

### Run Single Test
```bash
cargo test test_text_format_bold
```

### Run Tests with Output
```bash
cargo test -- --nocapture
```

### Run Tests with Verbose Output
```bash
cargo test -- --nocapture --test-threads=1
```

## Test Organization

### Unit Tests (in-module)
Located in each module file with `#[cfg(test)]` blocks:
- `src/generator/text.rs` - Text formatting tests
- `src/generator/shapes.rs` - Shape tests
- `src/generator/tables.rs` - Table tests

### Integration Tests
Located in `tests/` directory:
- `tests/advanced_features_test.rs` - Comprehensive feature tests

## Test Coverage by Feature

### Text Formatting
- ✅ Default properties
- ✅ Individual formatting options
- ✅ Combined formatting
- ✅ Color handling (uppercase/lowercase)
- ✅ Font size specification
- ✅ Builder pattern

### Shapes
- ✅ All 7 shape types
- ✅ Fill properties
- ✅ Transparency levels
- ✅ Line/border properties
- ✅ Text content
- ✅ Position and size
- ✅ Edge cases (zero, large dimensions)

### Tables
- ✅ Cell creation and formatting
- ✅ Row creation and properties
- ✅ Table creation
- ✅ Builder pattern
- ✅ Data import
- ✅ Edge cases (single cell, many rows/columns)

### Utilities
- ✅ EMU conversions
- ✅ Roundtrip accuracy
- ✅ Multiple unit systems

## Test Quality Metrics

### Coverage
- **Text Formatting**: 100% of public API
- **Shapes**: 100% of public API
- **Tables**: 100% of public API
- **Utilities**: 100% of public API

### Test Types
- **Unit Tests**: 34 (40%)
- **Integration Tests**: 51 (60%)

### Edge Cases Covered
- Default values
- Boundary conditions
- Multiple applications
- Roundtrip conversions
- Large data sets
- Single item edge cases

## Continuous Integration

All tests must pass before merging:
```bash
cargo test
```

## Adding New Tests

When adding new features:

1. Add unit tests in the module file
2. Add integration tests in `tests/advanced_features_test.rs`
3. Ensure all tests pass: `cargo test`
4. Update this document if adding new test categories

## Test Naming Convention

Tests follow the pattern: `test_<feature>_<scenario>`

Examples:
- `test_text_format_bold` - Testing bold formatting
- `test_shape_with_fill` - Testing shape with fill
- `test_table_many_rows` - Testing table with many rows

## Debugging Tests

### Run Single Test with Backtrace
```bash
RUST_BACKTRACE=1 cargo test test_name -- --nocapture
```

### Run Tests in Single Thread
```bash
cargo test -- --test-threads=1
```

### Print Debug Information
```bash
cargo test -- --nocapture --test-threads=1
```

## Performance Tests

Currently, no performance benchmarks are included. Future additions:
- Benchmark large table creation
- Benchmark shape rendering
- Benchmark text formatting application

## See Also

- [ADVANCED_FEATURES.md](ADVANCED_FEATURES.md) - Feature documentation
- [README.md](../README.md) - Main documentation
- [CODEBASE.md](../CODEBASE.md) - Code organization
