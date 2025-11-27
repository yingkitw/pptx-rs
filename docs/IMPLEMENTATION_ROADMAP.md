# Implementation Roadmap

This document outlines the current status and implementation plan for PowerPoint features.

## Current Status

### ✅ Implemented Features

**Text & Formatting**
- ✅ Slide titles and bullet points
- ✅ Font size customization (8pt to 96pt)
- ✅ Bold formatting (title and content)
- ✅ Special character escaping
- ✅ Unicode text support
- ✅ Long text handling

**Slide Structure**
- ✅ Single and multiple slides
- ✅ Slide layouts and masters
- ✅ Slide relationships
- ✅ Proper ZIP packaging

**CLI & Generation**
- ✅ CLI commands (create, from-md, info)
- ✅ Markdown to PPTX conversion
- ✅ PPTX file generation
- ✅ File validation

### ⏳ Defined But Not Implemented

**Tables**
- ⏳ Data structures defined (Table, TableRow, TableCell)
- ⏳ Builder API created
- ⏳ Tests written (12 tests)
- ❌ XML generation not implemented
- ❌ Not integrated into slides

**Charts**
- ⏳ Data structures defined (chart data types)
- ⏳ Tests written (7 tests)
- ❌ XML generation not implemented
- ❌ Not integrated into slides

**Images**
- ⏳ Metadata structures defined
- ⏳ Tests written (6 tests)
- ❌ Image embedding not implemented
- ❌ Not integrated into slides

**Advanced Text Styling**
- ⏳ TextFormat and FormattedText defined
- ⏳ Tests written (11 tests)
- ✅ Font size working
- ✅ Bold formatting working
- ❌ Italic not implemented
- ❌ Underline not implemented
- ❌ Color not implemented

## Implementation Priority

### Phase 1: Complete Text Styling (High Priority)
**Effort**: 2-3 hours
**Impact**: Improves text formatting capabilities

Tasks:
1. Add italic support to XML generation
2. Add underline support to XML generation
3. Add text color support to XML generation
4. Update tests and examples
5. Verify in PowerPoint

### Phase 2: Table Implementation (High Priority)
**Effort**: 4-6 hours
**Impact**: Enables data presentation

Tasks:
1. Design table XML structure
2. Implement table XML generation
3. Implement cell XML generation
4. Implement row XML generation
5. Handle table positioning and sizing
6. Add table to slide generation
7. Test with various table sizes
8. Verify in PowerPoint

### Phase 3: Image Implementation (Medium Priority)
**Effort**: 3-4 hours
**Impact**: Enables visual content

Tasks:
1. Design image embedding approach
2. Implement image file handling
3. Implement image XML generation
4. Handle image positioning and sizing
5. Add image to slide generation
6. Test with various image formats
7. Verify in PowerPoint

### Phase 4: Chart Implementation (Medium Priority)
**Effort**: 5-7 hours
**Impact**: Enables data visualization

Tasks:
1. Design chart XML structure
2. Implement chart XML generation
3. Implement chart data handling
4. Support multiple chart types (bar, pie, line)
5. Add chart to slide generation
6. Test with various data sets
7. Verify in PowerPoint

## Detailed Implementation Plans

### Text Styling Implementation

**Files to modify:**
- `src/generator/xml.rs` - Add italic, underline, color to XML
- `src/generator/text.rs` - Already has structures, needs XML integration

**XML Changes:**
```xml
<!-- Current (bold only) -->
<a:rPr lang="en-US" sz="4400" b="1"/>

<!-- Needed (with italic, underline, color) -->
<a:rPr lang="en-US" sz="4400" b="1" i="1" u="sng">
  <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
</a:rPr>
```

**Estimated effort**: 2-3 hours

### Table Implementation

**Files to create/modify:**
- `src/generator/tables_xml.rs` - New file for table XML generation
- `src/generator/mod.rs` - Export table XML functions
- `src/generator/builder.rs` - Integrate tables into slide generation

**XML Structure:**
```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>...</p:nvGraphicFramePr>
  <p:xfrm>...</p:xfrm>
  <a:graphic>
    <a:graphicData uri="...table...">
      <a:tbl>
        <a:tblPr>...</a:tblPr>
        <a:tblGrid>...</a:tblGrid>
        <a:tr>
          <a:tc>...</a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

**Estimated effort**: 4-6 hours

### Image Implementation

**Files to create/modify:**
- `src/generator/images.rs` - New file for image handling
- `src/generator/mod.rs` - Export image functions
- `src/generator/builder.rs` - Integrate images into slide generation

**XML Structure:**
```xml
<p:pic>
  <p:nvPicPr>...</p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId2"/>
    <a:stretch>...</a:stretch>
  </p:blipFill>
  <p:spPr>...</p:spPr>
</p:pic>
```

**Estimated effort**: 3-4 hours

### Chart Implementation

**Files to create/modify:**
- `src/generator/charts.rs` - New file for chart handling
- `src/generator/mod.rs` - Export chart functions
- `src/generator/builder.rs` - Integrate charts into slide generation

**XML Structure:**
```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>...</p:nvGraphicFramePr>
  <p:xfrm>...</p:xfrm>
  <a:graphic>
    <a:graphicData uri="...chart...">
      <c:chartSpace>
        <c:chart>
          <c:plotArea>
            <c:barChart>...</c:barChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

**Estimated effort**: 5-7 hours

## Implementation Strategy

### Step 1: Assess Current State
- [x] Identify what's implemented
- [x] Identify what's defined but not implemented
- [x] Create test coverage
- [x] Document structures

### Step 2: Implement Text Styling
1. Add italic XML attribute
2. Add underline XML attribute
3. Add color XML element
4. Update SlideContent to use TextFormat
5. Test and verify

### Step 3: Implement Tables
1. Design table XML structure
2. Create table XML generator
3. Integrate into slide generation
4. Test with various sizes
5. Verify in PowerPoint

### Step 4: Implement Images
1. Design image handling
2. Create image XML generator
3. Handle image embedding
4. Integrate into slide generation
5. Test with various formats

### Step 5: Implement Charts
1. Design chart XML structure
2. Create chart XML generator
3. Support multiple chart types
4. Integrate into slide generation
5. Test with various data

## Testing Strategy

For each feature:
1. Unit tests for XML generation
2. Integration tests for slide generation
3. PPTX file generation tests
4. Manual verification in PowerPoint

## Current Test Coverage

| Feature | Tests | Status |
|---------|-------|--------|
| Text formatting | 11 | ✅ Partial (bold/size only) |
| Tables | 12 | ⏳ Defined, not implemented |
| Charts | 7 | ⏳ Defined, not implemented |
| Images | 6 | ⏳ Defined, not implemented |
| Slides | 31 | ✅ Implemented |
| **Total** | **67** | **Mixed** |

## Known Limitations

### Current
- Text styling limited to bold and font size
- No italic, underline, or color support
- No table rendering
- No image embedding
- No chart generation

### By Design
- No VBA macro support (future)
- No animation support (future)
- No master slide customization (future)
- No theme customization (future)

## Next Steps

**Recommended order:**
1. Complete text styling (quick win)
2. Implement tables (high impact)
3. Implement images (visual content)
4. Implement charts (data visualization)

**Timeline estimate**: 14-20 hours for all features

## References

- PPTX XML Structure: `docs/ADVANCED_FEATURES.md`
- Test Coverage: `docs/ELEMENT_TESTS.md`, `docs/ADVANCED_ELEMENT_TESTS.md`
- Current Implementation: `src/generator/`
