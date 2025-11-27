# PPTX Alignment Testing

## Overview

This document describes the alignment testing framework for ensuring `ppt-rs` generates PPTX files that are compatible with the industry-standard `python-pptx` library.

## Purpose

Alignment testing ensures that:
- Generated PPTX files are valid and open correctly in PowerPoint
- Output structure matches python-pptx standards
- Compatibility is maintained across versions
- Regressions are caught early

## Quick Start

### 1. Generate Reference File (python-pptx)

```bash
python3 scripts/generate_reference.py
```

This creates `examples/output/reference_python_pptx.pptx` using python-pptx.

**Requirements:**
- Python 3.x
- python-pptx library: `pip install python-pptx`

### 2. Generate ppt-rs File

```bash
cargo run --example alignment_test
```

This creates `examples/output/alignment_test_ppt_rs.pptx` using ppt-rs.

### 3. Compare Files

```bash
python3 scripts/compare_pptx.py
```

Or with custom files:
```bash
python3 scripts/compare_pptx.py reference.pptx ppt_rs.pptx
```

## Alignment Score

The comparison script calculates an alignment score based on:
- **Matching XML files**: Percentage of XML files that match exactly
- **File structure**: ZIP archive structure comparison
- **File sizes**: Size differences (expected due to timestamps)

### Score Interpretation

- **95%+**: Excellent alignment - Minor formatting differences only
- **85-94%**: Good alignment - Some expected differences
- **<85%**: Poor alignment - Significant structural differences

## What Gets Compared

### Core Files (Must Match)
- `[Content_Types].xml` - Content type registry
- `_rels/.rels` - Package relationships
- `ppt/presentation.xml` - Presentation structure
- `docProps/core.xml` - Document metadata

### Expected Differences
- `ppt/theme/theme1.xml` - Theme definitions (formatting may differ)
- `ppt/viewProps.xml` - View properties (formatting may differ)
- File sizes - Timestamps and metadata cause size differences

## Alignment Status

**Current Status**: ⚠️ **In Progress**

### Completed
- ✅ Basic presentation structure
- ✅ Metadata (title, author, subject, keywords, comments)
- ✅ Slide content structure
- ✅ ZIP archive format

### In Progress
- ⚠️ Shape positioning alignment
- ⚠️ Text formatting alignment
- ⚠️ Theme and styling alignment

### Known Differences
- Theme XML structure (ppt-rs uses minimal theme)
- View properties formatting
- Some shape positioning may differ slightly

## Files

### Scripts
- `scripts/generate_reference.py` - Generate python-pptx reference
- `scripts/compare_pptx.py` - Compare PPTX files

### Examples
- `examples/alignment_test.rs` - Rust example for alignment testing

### Output
- `examples/output/reference_python_pptx.pptx` - Reference file
- `examples/output/alignment_test_ppt_rs.pptx` - ppt-rs output

## Troubleshooting

### python-pptx Not Installed
```bash
pip install python-pptx
```

### Files Not Found
Ensure you've run both generation steps:
1. `python3 scripts/generate_reference.py`
2. `cargo run --example alignment_test`

### Comparison Shows Many Differences
- Check that both files are valid PPTX
- Verify file paths are correct
- Review differences in detail to identify root causes

## Best Practices

1. **Run alignment tests regularly** - Catch regressions early
2. **Review differences carefully** - Not all differences are problems
3. **Document known differences** - Update this file as needed
4. **Aim for 95%+ alignment** - Excellent compatibility target

## Future Improvements

- [ ] Automated alignment testing in CI/CD
- [ ] More comprehensive test cases
- [ ] Visual comparison tools
- [ ] Alignment score tracking over time

---

*Last Updated: 2025-01-27*
*Alignment Score Target: 95%+*

