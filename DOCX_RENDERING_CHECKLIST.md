# DOCX/OOXML Rendering Checklist

A comprehensive checklist of elements and properties that need to be handled for accurate DOCX rendering.

## Table of Contents
1. [Document Structure](#1-document-structure)
2. [Text Formatting (Run Properties - w:rPr)](#2-text-formatting-run-properties---wrpr)
3. [Paragraph Formatting (w:pPr)](#3-paragraph-formatting-wppr)
4. [Styles System](#4-styles-system)
5. [Tables](#5-tables)
6. [Lists and Numbering](#6-lists-and-numbering)
7. [Images and Drawings](#7-images-and-drawings)
8. [Special Elements](#8-special-elements)
9. [Sections and Page Layout](#9-sections-and-page-layout)
10. [Headers and Footers](#10-headers-and-footers)
11. [Fields](#11-fields)
12. [Links and Navigation](#12-links-and-navigation)

---

## 1. Document Structure

### Package Structure
- [ ] ZIP container extraction
- [ ] `[Content_Types].xml` parsing
- [ ] `_rels/.rels` relationships
- [ ] `word/document.xml` (main content)
- [ ] `word/styles.xml` (styles definitions)
- [ ] `word/numbering.xml` (list definitions)
- [ ] `word/settings.xml` (document settings)
- [ ] `word/fontTable.xml` (font definitions)
- [ ] `word/theme/theme1.xml` (theme definitions)
- [ ] `word/_rels/document.xml.rels` (document relationships)
- [ ] `word/media/` (embedded images)
- [ ] `word/header*.xml` (headers)
- [ ] `word/footer*.xml` (footers)
- [ ] `docProps/core.xml` (document metadata)
- [ ] `docProps/app.xml` (application metadata)

### Document Body (w:body)
- [ ] Paragraphs (`w:p`)
- [ ] Tables (`w:tbl`)
- [ ] Section properties (`w:sectPr`)

---

## 2. Text Formatting (Run Properties - w:rPr)

### Basic Formatting
- [ ] **Bold** (`w:b`) - toggle property
- [ ] **Bold for complex script** (`w:bCs`)
- [ ] **Italic** (`w:i`) - toggle property
- [ ] **Italic for complex script** (`w:iCs`)
- [ ] **Underline** (`w:u`)
  - [ ] `val`: single, double, thick, dotted, dashed, dash, dotDash, dotDotDash, wave, wavyDouble, wavyHeavy, words, etc.
  - [ ] `color`: hex color (RRGGBB)
- [ ] **Strikethrough** (`w:strike`) - toggle property
- [ ] **Double strikethrough** (`w:dstrike`) - toggle property

### Font Properties
- [ ] **Font family** (`w:rFonts`)
  - [ ] `ascii`: ASCII characters (U+0000-U+007F)
  - [ ] `hAnsi`: High ANSI characters
  - [ ] `cs`: Complex script fonts
  - [ ] `eastAsia`: East Asian fonts
  - [ ] Theme font references (`asciiTheme`, `hAnsiTheme`, etc.)
- [ ] **Font size** (`w:sz`) - in half-points (1/144 inch)
- [ ] **Font size for complex script** (`w:szCs`)

### Color and Effects
- [ ] **Text color** (`w:color`)
  - [ ] `val`: hex color (RRGGBB) or "auto"
  - [ ] `themeColor`: theme color reference
  - [ ] `themeShade`: shade modifier
  - [ ] `themeTint`: tint modifier
- [ ] **Highlight/background color** (`w:highlight`)
  - [ ] Values: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black, white
- [ ] **Text shading** (`w:shd`)
  - [ ] `fill`: background color
  - [ ] `color`: pattern color
  - [ ] `val`: pattern type (clear, solid, pct10-pct90, diagStripe, etc.)

### Text Effects
- [ ] **All caps** (`w:caps`) - toggle property
- [ ] **Small caps** (`w:smallCaps`) - toggle property
- [ ] **Emboss** (`w:emboss`) - toggle property
- [ ] **Imprint/Engrave** (`w:imprint`) - toggle property
- [ ] **Outline** (`w:outline`) - toggle property
- [ ] **Shadow** (`w:shadow`) - toggle property
- [ ] **Hidden text** (`w:vanish`) - toggle property

### Position and Spacing
- [ ] **Superscript/Subscript** (`w:vertAlign`)
  - [ ] Values: baseline, superscript, subscript
- [ ] **Character spacing** (`w:spacing`) - in twips (1/1440 inch)
- [ ] **Character scale/width** (`w:w`) - percentage
- [ ] **Position** (`w:position`) - vertical position in half-points
- [ ] **Kerning** (`w:kern`) - minimum font size for kerning
- [ ] **Fit text** (`w:fitText`) - compress text to fit width

### Text Borders
- [ ] **Character border** (`w:bdr`)
  - [ ] `val`: border style
  - [ ] `sz`: border width (eighths of a point)
  - [ ] `space`: padding
  - [ ] `color`: border color

### Complex Script Properties
- [ ] **Complex script** (`w:cs`) - mark as complex script
- [ ] **Right-to-left** (`w:rtl`)

### Character Style Reference
- [ ] **Style reference** (`w:rStyle`)

---

## 3. Paragraph Formatting (w:pPr)

### Alignment
- [ ] **Horizontal alignment** (`w:jc`)
  - [ ] Values: left, center, right, both (justify), distribute

### Indentation (`w:ind`)
- [ ] `left`/`start`: left indentation (twips)
- [ ] `right`/`end`: right indentation (twips)
- [ ] `firstLine`: first line additional indent
- [ ] `hanging`: hanging indent (negative first line)
- [ ] `leftChars`/`rightChars`: character-based indentation

### Spacing (`w:spacing`)
- [ ] `before`: space before paragraph (twips)
- [ ] `after`: space after paragraph (twips)
- [ ] `beforeLines`/`afterLines`: line-based spacing
- [ ] `line`: line spacing value
- [ ] `lineRule`: auto, exact, atLeast
- [ ] `beforeAutospacing`/`afterAutospacing`: automatic spacing

### Borders (`w:pBdr`)
- [ ] `top`: top border
- [ ] `bottom`: bottom border
- [ ] `left`: left border
- [ ] `right`: right border
- [ ] `between`: border between identical paragraphs
- [ ] `bar`: bar border (for facing pages)
- [ ] Border attributes:
  - [ ] `val`: border style (single, double, dotted, dashed, wave, etc.)
  - [ ] `sz`: width in eighths of a point
  - [ ] `space`: spacing from text (points)
  - [ ] `color`: hex color
  - [ ] `shadow`: shadow effect

### Shading (`w:shd`)
- [ ] `fill`: background color
- [ ] `color`: pattern color
- [ ] `val`: pattern type

### Tabs (`w:tabs`)
- [ ] **Tab stops** (`w:tab`)
  - [ ] `val`: left, center, right, decimal, bar, clear
  - [ ] `pos`: position in twips
  - [ ] `leader`: none, dot, hyphen, underscore, heavy, middleDot

### Text Flow Control
- [ ] **Keep lines together** (`w:keepLines`)
- [ ] **Keep with next** (`w:keepNext`)
- [ ] **Page break before** (`w:pageBreakBefore`)
- [ ] **Widow/orphan control** (`w:widowControl`)
- [ ] **Suppress line numbers** (`w:suppressLineNumbers`)
- [ ] **Suppress auto-hyphens** (`w:suppressAutoHyphens`)

### Text Direction
- [ ] **Bidirectional** (`w:bidi`)
- [ ] **Text direction** (`w:textDirection`)
- [ ] **Text alignment** (`w:textAlignment`)
  - [ ] Values: auto, baseline, bottom, center, top

### Outline Level
- [ ] **Outline level** (`w:outlineLvl`) - 0-9 for TOC

### Paragraph Style
- [ ] **Style reference** (`w:pStyle`)

### Numbering Reference
- [ ] **Numbering properties** (`w:numPr`)
  - [ ] `ilvl`: list level
  - [ ] `numId`: numbering definition ID

### Section Properties in Paragraph
- [ ] **Section properties** (`w:sectPr`) - for section breaks

### Run Properties for Paragraph Mark
- [ ] **Paragraph mark formatting** (`w:rPr` within `w:pPr`)

### Frame Properties
- [ ] **Text frame** (`w:framePr`)
  - [ ] Positioning attributes
  - [ ] Wrapping attributes

### Contextual Spacing
- [ ] **Contextual spacing** (`w:contextualSpacing`)

---

## 4. Styles System

### Style Types (`w:style`)
- [ ] **Paragraph styles** (`w:type="paragraph"`)
- [ ] **Character styles** (`w:type="character"`)
- [ ] **Table styles** (`w:type="table"`)
- [ ] **Numbering styles** (`w:type="numbering"`)

### Style Properties
- [ ] `w:styleId`: unique identifier
- [ ] `w:name`: display name
- [ ] `w:basedOn`: parent style inheritance
- [ ] `w:next`: following paragraph style
- [ ] `w:link`: linked style (paragraph-character link)
- [ ] `w:default`: default style marker
- [ ] `w:qFormat`: quick format style
- [ ] `w:uiPriority`: sort order in UI
- [ ] `w:semiHidden`/`w:hidden`: visibility
- [ ] `w:unhideWhenUsed`: auto-show when used

### Style Inheritance Hierarchy
1. Document defaults (`w:docDefaults`)
2. Table styles
3. Numbering styles
4. Paragraph styles (including inherited `w:rPr`)
5. Character/run styles
6. Direct formatting

### Toggle Properties Behavior
- [ ] Bold, italic, caps, smallCaps, strike, dstrike, outline, shadow, emboss, imprint, vanish
- [ ] Toggle through style hierarchy (odd count = true)

### Document Defaults (`w:docDefaults`)
- [ ] `w:rPrDefault`: default run properties
- [ ] `w:pPrDefault`: default paragraph properties

### Theme Colors and Fonts
- [ ] Theme color references
- [ ] Theme font references (major/minor)
- [ ] Color scheme mapping (`w:clrSchemeMapping`)

---

## 5. Tables

### Table Structure (`w:tbl`)
- [ ] Table properties (`w:tblPr`)
- [ ] Table grid (`w:tblGrid`)
- [ ] Table rows (`w:tr`)

### Table Properties (`w:tblPr`)
- [ ] **Style** (`w:tblStyle`)
- [ ] **Width** (`w:tblW`)
  - [ ] `type`: auto, dxa (twips), pct (percent), nil
  - [ ] `w`: width value
- [ ] **Alignment** (`w:jc`): left, center, right
- [ ] **Indentation** (`w:tblInd`)
- [ ] **Borders** (`w:tblBorders`)
  - [ ] `top`, `bottom`, `left`, `right`
  - [ ] `insideH`, `insideV`
- [ ] **Cell margins** (`w:tblCellMar`)
  - [ ] `top`, `bottom`, `left`/`start`, `right`/`end`
- [ ] **Cell spacing** (`w:tblCellSpacing`)
- [ ] **Layout** (`w:tblLayout`): autofit, fixed
- [ ] **Table look** (`w:tblLook`) - conditional formatting flags
  - [ ] `firstRow`, `lastRow`, `firstColumn`, `lastColumn`
  - [ ] `noHBand`, `noVBand`
- [ ] **Floating table** (`w:tblpPr`)
  - [ ] Position attributes
  - [ ] Distance from text

### Table Grid (`w:tblGrid`)
- [ ] **Grid columns** (`w:gridCol`)
  - [ ] `w:w`: column width in twips

### Row Properties (`w:trPr`)
- [ ] **Row height** (`w:trHeight`)
  - [ ] `val`: height value
  - [ ] `hRule`: auto, atLeast, exact
- [ ] **Header row** (`w:tblHeader`) - repeat on page break
- [ ] **Can split** (`w:cantSplit`) - prevent row split
- [ ] **Justify** (`w:jc`)

### Cell Properties (`w:tcPr`)
- [ ] **Width** (`w:tcW`)
- [ ] **Horizontal span** (`w:gridSpan`)
- [ ] **Vertical merge** (`w:vMerge`)
  - [ ] `val`: restart, continue
- [ ] **Horizontal merge** (`w:hMerge`) - legacy
- [ ] **Borders** (`w:tcBorders`)
  - [ ] `top`, `bottom`, `left`/`start`, `right`/`end`
  - [ ] `insideH`, `insideV`
  - [ ] `tl2br`, `tr2bl` (diagonal)
- [ ] **Shading** (`w:shd`)
- [ ] **Cell margins** (`w:tcMar`)
- [ ] **Vertical alignment** (`w:vAlign`): top, center, bottom
- [ ] **Text direction** (`w:textDirection`)
- [ ] **Fit text** (`w:tcFitText`)
- [ ] **No wrap** (`w:noWrap`)
- [ ] **Hide mark** (`w:hideMark`)

### Border Conflict Resolution
- [ ] Table border vs cell border
- [ ] Overlapping cell borders
- [ ] Priority: direct > style
- [ ] Wider border wins when equal priority

---

## 6. Lists and Numbering

### Numbering Definitions (`numbering.xml`)

#### Abstract Numbering (`w:abstractNum`)
- [ ] `w:abstractNumId`: unique ID
- [ ] `w:multiLevelType`: singleLevel, multilevel, hybridMultilevel
- [ ] `w:nsid`: number scheme ID

#### Level Definitions (`w:lvl`)
- [ ] `w:ilvl`: level index (0-8)
- [ ] **Start value** (`w:start`)
- [ ] **Number format** (`w:numFmt`)
  - [ ] Values: decimal, upperRoman, lowerRoman, upperLetter, lowerLetter, bullet, none, ordinal, cardinalText, ordinalText, etc.
- [ ] **Level text** (`w:lvlText`)
  - [ ] `val`: text pattern (e.g., "%1.", "%1.%2")
- [ ] **Justification** (`w:lvlJc`): left, center, right
- [ ] **Paragraph properties** (`w:pPr`)
  - [ ] Indentation for list level
  - [ ] Tab stops
- [ ] **Run properties** (`w:rPr`)
  - [ ] Font for bullet/number
- [ ] **Picture bullet** (`w:lvlPicBulletId`)
- [ ] **Suffix** (`w:suff`): tab, space, nothing
- [ ] **Restart** (`w:lvlRestart`)
- [ ] **Legal numbering** (`w:isLgl`)
- [ ] **Style link** (`w:pStyle`)

#### Numbering Instance (`w:num`)
- [ ] `w:numId`: instance ID
- [ ] **Abstract reference** (`w:abstractNumId`)
- [ ] **Level overrides** (`w:lvlOverride`)
  - [ ] `ilvl`: level to override
  - [ ] Override properties

### Picture Bullets
- [ ] `w:numPicBullet`
- [ ] Reference to embedded image

---

## 7. Images and Drawings

### Inline Images (`w:drawing` > `wp:inline`)
- [ ] **Extent** (`wp:extent`)
  - [ ] `cx`, `cy`: size in EMUs (914400 EMUs = 1 inch)
- [ ] **Effect extent** (`wp:effectExtent`)
- [ ] **Document properties** (`wp:docPr`)
  - [ ] `id`, `name`, `descr`
- [ ] **Graphic frame locks** (`wp:cNvGraphicFramePr`)
  - [ ] `noChangeAspect`
- [ ] **Graphic data** (`a:graphic` > `a:graphicData`)

### Floating Images (`w:drawing` > `wp:anchor`)
- [ ] **Position** (`wp:positionH`, `wp:positionV`)
  - [ ] `relativeFrom`: page, margin, column, paragraph, character, line
  - [ ] `posOffset`, `align`, `percentOffset`
- [ ] **Simple position** (`wp:simplePos`)
- [ ] **Extent** (`wp:extent`)
- [ ] **Effect extent** (`wp:effectExtent`)
- [ ] **Text wrapping**
  - [ ] `wp:wrapNone`
  - [ ] `wp:wrapSquare`
  - [ ] `wp:wrapTight`
  - [ ] `wp:wrapThrough`
  - [ ] `wp:wrapTopAndBottom`
- [ ] **Wrap attributes**
  - [ ] `wrapText`: bothSides, left, right, largest
  - [ ] `distT`, `distB`, `distL`, `distR`
- [ ] **Behind/in front of text** (`behindDoc`)
- [ ] **Relative height** (`relativeHeight`)
- [ ] **Lock anchor** (`locked`)
- [ ] **Layout in cell** (`layoutInCell`)
- [ ] **Allow overlap** (`allowOverlap`)

### Picture Definition (`pic:pic`)
- [ ] **Non-visual properties** (`pic:nvPicPr`)
  - [ ] `pic:cNvPr`: id, name, description
  - [ ] `pic:cNvPicPr`: picture locks
- [ ] **Blob fill** (`pic:blipFill`)
  - [ ] `a:blip`: image reference (`r:embed`, `r:link`)
  - [ ] `a:stretch` / `a:tile`
  - [ ] `a:srcRect`: cropping
- [ ] **Shape properties** (`pic:spPr`)
  - [ ] Transform (`a:xfrm`)
  - [ ] Geometry (`a:prstGeom`)
  - [ ] Effects

### VML Images (legacy)
- [ ] `w:pict` container
- [ ] `v:shape`, `v:imagedata`

### Image Formats
- [ ] PNG
- [ ] JPEG
- [ ] GIF
- [ ] TIFF
- [ ] EMF/WMF (vector)
- [ ] SVG (newer versions)

---

## 8. Special Elements

### Breaks (`w:br`)
- [ ] **Line break** (default, `type="textWrapping"`)
- [ ] **Page break** (`type="page"`)
- [ ] **Column break** (`type="column"`)
- [ ] **Clear** attribute: none, left, right, all

### Special Characters
- [ ] **Tab** (`w:tab`)
- [ ] **Carriage return** (`w:cr`)
- [ ] **Soft hyphen** (`w:softHyphen`)
- [ ] **Non-breaking hyphen** (`w:noBreakHyphen`)
- [ ] **Symbol** (`w:sym`)
  - [ ] `char`: character code
  - [ ] `font`: font name

### Comments
- [ ] **Comment reference** (`w:commentReference`)
- [ ] **Comment range start** (`w:commentRangeStart`)
- [ ] **Comment range end** (`w:commentRangeEnd`)
- [ ] **Comment content** (`w:comments` in comments.xml)

### Footnotes and Endnotes
- [ ] **Footnote reference** (`w:footnoteReference`)
- [ ] **Endnote reference** (`w:endnoteReference`)
- [ ] **Separator marks** (`w:separator`, `w:continuationSeparator`)
- [ ] Content in `footnotes.xml` / `endnotes.xml`

### Track Changes
- [ ] **Insertions** (`w:ins`)
- [ ] **Deletions** (`w:del`)
- [ ] **Move from/to** (`w:moveFrom`, `w:moveTo`)
- [ ] **Format changes** (`w:rPrChange`, `w:pPrChange`)
- [ ] Revision attributes: `author`, `date`, `id`

### Text Boxes
- [ ] **DrawingML text boxes** (`wps:wsp`)
- [ ] **VML text boxes** (`v:textbox`)
- [ ] Text frame paragraphs (`w:framePr`)

### Shapes
- [ ] **DrawingML shapes** (`wps:wsp`)
  - [ ] Preset geometry (`a:prstGeom`)
  - [ ] Custom geometry (`a:custGeom`)
  - [ ] Fill (solid, gradient, pattern, picture)
  - [ ] Outline
  - [ ] Effects
- [ ] **VML shapes** (legacy)

### Math/Equations
- [ ] Office Math ML (`m:oMath`)
- [ ] OLE objects for equations

---

## 9. Sections and Page Layout

### Section Properties (`w:sectPr`)
- [ ] **Page size** (`w:pgSz`)
  - [ ] `w:w`, `w:h`: dimensions in twips
  - [ ] `w:orient`: portrait, landscape
- [ ] **Page margins** (`w:pgMar`)
  - [ ] `top`, `bottom`, `left`/`start`, `right`/`end`
  - [ ] `header`, `footer`: distance from edge
  - [ ] `gutter`: gutter margin
- [ ] **Columns** (`w:cols`)
  - [ ] `num`: number of columns
  - [ ] `space`: spacing between
  - [ ] `sep`: separator line
  - [ ] Individual column definitions (`w:col`)
- [ ] **Page borders** (`w:pgBorders`)
  - [ ] `top`, `bottom`, `left`, `right`
  - [ ] `offsetFrom`: page, text
  - [ ] `display`: allPages, firstPage, notFirstPage
- [ ] **Line numbers** (`w:lnNumType`)
  - [ ] `countBy`: interval
  - [ ] `start`: starting number
  - [ ] `restart`: continuous, newPage, newSection
  - [ ] `distance`: from text
- [ ] **Page numbering** (`w:pgNumType`)
  - [ ] `fmt`: format (decimal, upperRoman, etc.)
  - [ ] `start`: starting number
- [ ] **Section type** (`w:type`)
  - [ ] Values: continuous, nextPage, evenPage, oddPage, nextColumn
- [ ] **Vertical alignment** (`w:vAlign`)
- [ ] **Text direction** (`w:textDirection`)
- [ ] **Document grid** (`w:docGrid`)
- [ ] **Paper source** (`w:paperSrc`)
- [ ] **Form protection** (`w:formProt`)
- [ ] **Title page** (`w:titlePg`)

### Header/Footer References
- [ ] **Header reference** (`w:headerReference`)
  - [ ] `type`: default, first, even
  - [ ] `r:id`: relationship ID
- [ ] **Footer reference** (`w:footerReference`)
  - [ ] `type`: default, first, even
  - [ ] `r:id`: relationship ID

---

## 10. Headers and Footers

### Header Parts (`header*.xml`)
- [ ] Root element: `w:hdr`
- [ ] Contains same block-level content as body
- [ ] Multiple headers per section (default, first, even)

### Footer Parts (`footer*.xml`)
- [ ] Root element: `w:ftr`
- [ ] Contains same block-level content as body
- [ ] Multiple footers per section (default, first, even)

### Header/Footer Content
- [ ] Paragraphs
- [ ] Tables
- [ ] Images
- [ ] Fields (especially page numbers)
- [ ] Tab stops for positioning

---

## 11. Fields

### Simple Fields (`w:fldSimple`)
- [ ] `instr`: field instruction
- [ ] Cached result as content

### Complex Fields
- [ ] **Field start** (`w:fldChar fldCharType="begin"`)
- [ ] **Instructions** (`w:instrText`)
- [ ] **Separator** (`w:fldChar fldCharType="separate"`)
- [ ] **Result** (runs with cached value)
- [ ] **Field end** (`w:fldChar fldCharType="end"`)

### Common Field Types
- [ ] **PAGE**: current page number
- [ ] **NUMPAGES**: total pages
- [ ] **DATE**: current date
- [ ] **TIME**: current time
- [ ] **AUTHOR**: document author
- [ ] **TITLE**: document title
- [ ] **FILENAME**: file name
- [ ] **FILESIZE**: file size
- [ ] **TOC**: table of contents
- [ ] **REF**: cross-reference
- [ ] **SEQ**: sequence number
- [ ] **HYPERLINK**: hyperlink
- [ ] **MERGEFIELD**: mail merge
- [ ] **IF**: conditional

### Field Switches
- [ ] General formatting (`\*`)
- [ ] Date/time formatting (`\@`)
- [ ] Numeric formatting (`\#`)

---

## 12. Links and Navigation

### Hyperlinks (`w:hyperlink`)
- [ ] **External links**
  - [ ] `r:id`: relationship to external URL
  - [ ] `w:docLocation`: location within target
- [ ] **Internal links**
  - [ ] `w:anchor`: bookmark name
- [ ] **Attributes**
  - [ ] `w:tgtFrame`: target frame
  - [ ] `w:tooltip`: tooltip text
  - [ ] `w:history`: add to history

### Bookmarks
- [ ] **Bookmark start** (`w:bookmarkStart`)
  - [ ] `w:id`: unique ID
  - [ ] `w:name`: bookmark name
- [ ] **Bookmark end** (`w:bookmarkEnd`)
  - [ ] `w:id`: matching ID

### Table of Contents
- [ ] TOC field with switches
- [ ] Entry styles (`TOC1`, `TOC2`, etc.)
- [ ] Hyperlinks to headings

---

## Measurement Units Reference

| Unit | Description | Conversion |
|------|-------------|------------|
| Twip | 1/20 of a point | 1440 twips = 1 inch |
| Half-point | 1/2 of a point | 2 half-points = 1 point |
| Eighth-point | 1/8 of a point | 8 = 1 point (borders) |
| EMU | English Metric Unit | 914400 EMUs = 1 inch |
| Percent | Percentage | 100% = full |

---

## Priority Order for Rendering

### High Priority (Core Functionality)
1. Basic text rendering
2. Font family, size, color
3. Bold, italic, underline
4. Paragraph alignment
5. Basic spacing (before, after, line)
6. Simple tables
7. Inline images
8. Basic lists

### Medium Priority
9. All underline styles
10. All border styles
11. Complex table features (merge, borders)
12. Floating images
13. Multi-level lists
14. Styles inheritance
15. Headers/footers
16. Page breaks
17. Tabs

### Lower Priority
18. Comments
19. Footnotes/endnotes
20. Track changes
21. Fields (dynamic)
22. Text boxes
23. Shapes
24. Advanced effects (shadow, emboss)
25. Math equations
26. Complex page layouts

---

## Implementation Notes

### Toggle Properties
Toggle properties (bold, italic, etc.) have special inheritance rules:
- If explicitly set in direct formatting, use that value
- Through style hierarchy, first encountered value wins
- Multiple levels: XOR behavior (odd count = true)

### Color Resolution
1. Check explicit color value
2. Check theme color reference
3. Apply shade/tint modifiers
4. Resolve "auto" based on context

### Border Conflict Resolution
1. Wider border wins
2. If same width, darker color wins
3. Direct formatting > table style > cell style

### Whitespace Handling
- `xml:space="preserve"` must be respected
- Leading/trailing spaces in `w:t` elements
- Tab characters (`w:tab`)
