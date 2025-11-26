# Test-Generated PPTX Files

This document describes the PPTX files automatically generated during `cargo test`.

## Overview

When you run `cargo test`, the test suite not only validates functionality but also generates example PPTX files for manual verification. These files are created in `target/test_output/` and demonstrate various features and capabilities.

## Generated Files

### 1. test_text_formatting.pptx
**Purpose**: Demonstrates text formatting capabilities
**Slides**: 3
**Features**:
- Custom font sizes (52pt title, 24pt content)
- Bold formatting on both title and content
- Mixed formatting combinations
- Variable font sizes across slides

**Verification**:
```bash
unzip -p target/test_output/test_text_formatting.pptx ppt/slides/slide1.xml | grep -o 'sz="[0-9]*"'
```

### 2. test_simple.pptx
**Purpose**: Basic presentation with standard formatting
**Slides**: 3
**Features**:
- Introduction slide
- Key points slide
- Conclusion slide
- Default formatting (44pt title, 28pt content)

**Use Case**: Baseline reference for standard presentations

### 3. test_large_presentation.pptx
**Purpose**: Demonstrates scaling with multiple slides
**Slides**: 5
**Features**:
- Progressive font size increases
- Consistent content structure
- Varied title sizes (44pt to 52pt)
- Multiple bullet points per slide

**Use Case**: Testing larger presentations with varying styles

### 4. test_styled_presentation.pptx
**Purpose**: Comprehensive styling demonstration
**Slides**: 3
**Features**:
- Large title (60pt) with small content (16pt)
- Regular title (44pt) with standard content (28pt)
- Bold emphasis (48pt title, 26pt content, bold)
- Different layout densities

**Verification**:
```bash
unzip -p target/test_output/test_styled_presentation.pptx ppt/slides/slide1.xml | grep -o 'sz="[0-9]*"'
```

### 5. test_bold_content.pptx
**Purpose**: Demonstrates bold formatting
**Slides**: 1
**Features**:
- Bold title (true)
- Bold content (true)
- All text emphasized with bold

**Verification**:
```bash
unzip -p target/test_output/test_bold_content.pptx ppt/slides/slide1.xml | grep -o 'b="[01]"'
```

### 6. test_mixed_styles.pptx
**Purpose**: Shows mixed formatting combinations
**Slides**: 2
**Features**:
- Slide 1: 48pt title, 24pt content
- Slide 2: 44pt title, 28pt content
- Different formatting per slide

**Use Case**: Testing varied formatting across slides

## File Specifications

All generated files:
- ✅ Valid Microsoft PowerPoint 2007+ format
- ✅ Proper ZIP structure
- ✅ Valid XML components
- ✅ Readable in PowerPoint, LibreOffice, Google Slides
- ✅ ECMA-376 compliant

## Accessing Generated Files

### Location
```
target/test_output/
```

### Running Tests
```bash
# Generate all test files
cargo test

# Generate only specific test
cargo test test_generate_text_formatting_pptx
```

### Viewing Files
```bash
# List all generated files
ls -lh target/test_output/*.pptx

# Open in PowerPoint (macOS)
open target/test_output/test_text_formatting.pptx

# Open in LibreOffice
libreoffice target/test_output/test_text_formatting.pptx
```

## Verifying Features

### Check Font Sizes
```bash
unzip -p target/test_output/test_text_formatting.pptx ppt/slides/slide1.xml | grep -o 'sz="[0-9]*"'
```

### Check Bold Formatting
```bash
unzip -p target/test_output/test_bold_content.pptx ppt/slides/slide1.xml | grep -o 'b="[01]"'
```

### Check File Format
```bash
file target/test_output/test_text_formatting.pptx
```

### Verify ZIP Structure
```bash
unzip -l target/test_output/test_text_formatting.pptx | head -20
```

## Test Coverage

The generated files cover:

| Feature | File | Verification |
|---------|------|--------------|
| Font Sizes | test_text_formatting.pptx | Multiple sz values |
| Bold Text | test_bold_content.pptx | b="1" attributes |
| Mixed Styles | test_mixed_styles.pptx | Varied formatting |
| Large Presentation | test_large_presentation.pptx | 5 slides |
| Styled Content | test_styled_presentation.pptx | 3 different styles |
| Simple Baseline | test_simple.pptx | Standard formatting |

## Manual Verification Steps

1. **Run tests**:
   ```bash
   cargo test
   ```

2. **Open a generated file**:
   ```bash
   open target/test_output/test_text_formatting.pptx
   ```

3. **Verify in PowerPoint**:
   - Check title font size (should be 52pt)
   - Check content font size (should be 24pt)
   - Verify text is properly formatted

4. **Verify in LibreOffice**:
   - Open the file
   - Check formatting is preserved
   - Verify no corruption

## Troubleshooting

### Files Not Generated
- Ensure tests passed: `cargo test`
- Check target/test_output exists: `ls target/test_output/`
- Verify disk space is available

### Files Corrupted
- Delete and regenerate: `rm -rf target/test_output && cargo test`
- Check for compilation errors: `cargo build`

### Font Sizes Not Visible
- Open in PowerPoint or LibreOffice
- Check View menu for zoom level
- Verify file is not corrupted: `file target/test_output/test_text_formatting.pptx`

## Continuous Integration

These files are generated during:
- Local testing: `cargo test`
- CI/CD pipelines
- Pre-commit hooks

They serve as:
- Regression detection
- Feature validation
- Manual verification
- Documentation examples

## Cleanup

To remove generated test files:
```bash
rm -rf target/test_output/
```

To regenerate:
```bash
cargo test
```

## See Also

- [TESTING.md](TESTING.md) - Test suite documentation
- [ADVANCED_FEATURES.md](ADVANCED_FEATURES.md) - Feature documentation
- [README.md](../README.md) - Main documentation
