import 'package:collection/collection.dart';
import 'package:xml/xml.dart';

/// Snapshot of the top-level body structure of a DOCX, addressable by stable
/// 0-based indices for rewriting.
///
/// Only top-level body elements are indexed: paragraphs (`<w:p>`) and tables
/// (`<w:tbl>`) appearing directly under `<w:body>`. Nested content (a paragraph
/// inside a table cell) is not in [paragraphs] — table cells are addressable
/// via [tables] indices.
class DocxEditableStructure {
  DocxEditableStructure({
    required this.paragraphs,
    required this.tables,
  });

  final List<DocxParagraphInfo> paragraphs;
  final List<DocxTableInfo> tables;
}

class DocxParagraphInfo {
  DocxParagraphInfo({required this.pIdx, required this.text});

  /// 0-based index of this paragraph among top-level body paragraphs in
  /// document order.
  final int pIdx;

  /// Concatenated text content of all `<w:t>` runs in the paragraph.
  final String text;
}

class DocxTableInfo {
  DocxTableInfo({required this.tIdx, required this.rows});

  /// 0-based index of this table among top-level body tables in document order.
  final int tIdx;

  /// Cell text per row: `rows[r][c]` is the concatenated text of cell `c` in
  /// row `r`.
  final List<List<String>> rows;
}

/// One cell in a [TemplatedRow]. The cell will be written as a paragraph
/// containing a single SDT with the given tag/alias and placeholder text.
class TemplatedCell {
  /// If [alias] is omitted it defaults to [tag]. The existing fill pipeline
  /// keys lookups off the alias, so for tags containing `/` the alias must
  /// match the full path verbatim or the tag won't resolve at fill time.
  ///
  /// Set [image] to `true` for cells whose binding resolves to an image
  /// (e.g. `signee/signature`). [DocxTemplate.rewriteTableRows] then
  /// emits an `ImgView`-classified SDT with a placeholder `<w:drawing>`
  /// for that cell, mirroring [DocxTemplate.replaceCellContentWithImage]
  /// for non-table-row cells.
  const TemplatedCell({
    required this.tag,
    this.alias,
    this.placeholder,
    this.image = false,
    this.imageWidthEmu = 1828800,
    this.imageHeightEmu = 457200,
  });

  final String tag;
  final String? alias;
  final String? placeholder;
  final bool image;
  final int imageWidthEmu;
  final int imageHeightEmu;

  String get effectiveAlias => alias ?? tag;
}

/// Recipe for the single data row inserted by [DocxTemplate.rewriteTableRows].
///
/// The row is wrapped in an outer SDT (`wrapperTag` / `wrapperAlias`) so that
/// the existing template engine recognises it as a table tag and repeats the
/// row per content item at fill time. If [wrapperAlias] is omitted it
/// defaults to [wrapperTag] — the fill pipeline keys lookups off the alias,
/// so they must match for paths containing `/`.
class TemplatedRow {
  TemplatedRow({
    required this.wrapperTag,
    String? wrapperAlias,
    required this.cells,
  }) : wrapperAlias = wrapperAlias ?? wrapperTag;

  final String wrapperTag;
  final String wrapperAlias;
  final List<TemplatedCell> cells;
}

/// Allocates monotonically increasing SDT id values within an editing session.
/// Negative ids match the convention used by Word for generated SDTs.
class SdtIdAllocator {
  int _next = -1000000;

  String next() {
    final v = _next.toString();
    _next++;
    return v;
  }
}

/// Builds a `<w:sdt>` element wrapping [contentChildren] with the given tag
/// metadata. The `w` namespace prefix is expected to be bound at the document
/// root (it always is in valid DOCX files).
XmlElement buildSdt({
  required String tag,
  required String alias,
  required String id,
  required List<XmlNode> contentChildren,
}) {
  XmlName w(String local) => XmlName(local, 'w');

  XmlElement valEl(String local, String val) => XmlElement(
        w(local),
        [XmlAttribute(w('val'), val)],
      );

  final sdtPr = XmlElement(w('sdtPr'), [], [
    valEl('alias', alias),
    valEl('tag', tag),
    valEl('id', id),
  ]);

  final sdtContent = XmlElement(w('sdtContent'), [], contentChildren);

  return XmlElement(w('sdt'), [], [sdtPr, sdtContent]);
}

/// Builds a `<w:r><w:t xml:space="preserve">text</w:t></w:r>` run.
XmlElement buildRun({required String text, XmlElement? rPr}) {
  XmlName w(String local) => XmlName(local, 'w');
  final t = XmlElement(
    w('t'),
    [XmlAttribute(XmlName('space', 'xml'), 'preserve')],
    [XmlText(text)],
  );
  return XmlElement(w('r'), [], [
    if (rPr != null) rPr,
    t,
  ]);
}

/// Returns a deep copy of the first `<w:rPr>` found inside any `<w:r>` that
/// descends from [scope], or null when the scope contains no run with run
/// properties. Used to inherit the surrounding run's font/size/colour onto
/// freshly inserted SDT placeholder runs so backfill fields render in the
/// document's natural style instead of Word's default Calibri 11.
XmlElement? findFirstRunRpr(XmlElement scope) {
  for (final descendant in scope.descendants) {
    if (descendant is! XmlElement) continue;
    if (descendant.name.local != 'r') continue;
    final rPr = descendant.children
        .whereType<XmlElement>()
        .firstWhereOrNull((e) => e.name.local == 'rPr');
    if (rPr != null) return rPr.copy();
  }
  return null;
}

/// Returns an rPr that inherits everything from [base] (or starts empty if
/// null) and overrides the `<w:color>` child with [hexColour] (e.g.
/// `0070C0`). Hex value should be 6 hex chars without `#`. Pass null to
/// skip the colour override entirely — the original colour is preserved.
XmlElement? rPrWithColor({XmlElement? base, String? hexColour}) {
  if (hexColour == null) return base;
  XmlName w(String local) => XmlName(local, 'w');

  final clone = base != null
      ? base.copy()
      : XmlElement(w('rPr'), [], []);
  // Strip any existing color child so we don't end up with two.
  clone.children.removeWhere(
    (n) => n is XmlElement && n.name.local == 'color',
  );
  // Insert color near the front for readability — Word doesn't care about
  // child order inside rPr but it makes diffs easier to scan.
  clone.children.insert(
    0,
    XmlElement(w('color'), [XmlAttribute(w('val'), hexColour)]),
  );
  return clone;
}

/// Builds a `<w:p>` containing a single run with [text]. Used when generating
/// fresh paragraphs (e.g. for templated table cells).
XmlElement buildParagraph({required String text}) {
  XmlName w(String local) => XmlName(local, 'w');
  return XmlElement(w('p'), [], [buildRun(text: text)]);
}

/// 1×1 transparent PNG embedded as a placeholder image. Inserted into
/// `word/media/` whenever an image SDT is created via
/// [DocxTemplate.replaceCellContentWithImage]. The fill pipeline's
/// [ImgView] swaps this placeholder for the real bytes at fill time, so
/// what's actually here doesn't matter — only that there's *something*
/// for Word to attach the SDT to.
const List<int> kPlaceholderPngBytes = [
  0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG header
  0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
  0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89,
  0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, // IDAT
  0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05,
  0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4,
  0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND
  0xAE, 0x42, 0x60, 0x82,
];

/// Builds a `<w:sdt>` whose content is an inline `<w:drawing>` referencing
/// the relationship id [relId]. The fill pipeline classifies this as an
/// `ImgView` (because the SDT tag is "img") and swaps the underlying
/// `r:embed` for the real image bytes at fill time.
///
/// Sizes are in EMU (914400 EMU = 1 inch). Defaults give a roughly
/// 2 in × 0.5 in signature box that looks reasonable inside a Word table
/// cell.
XmlElement buildImageSdt({
  required String tag,
  required String alias,
  required String id,
  required String relId,
  required int widthEmu,
  required int heightEmu,
  required int docPrId,
}) {
  XmlName w(String local) => XmlName(local, 'w');

  XmlElement valEl(String local, String val) => XmlElement(
        w(local),
        [XmlAttribute(w('val'), val)],
      );

  final sdtPr = XmlElement(w('sdtPr'), [], [
    valEl('alias', alias),
    valEl('tag', tag),
    valEl('id', id),
  ]);

  // Build the namespaced drawing tree. Word is fussy here — namespace
  // prefixes must match the schema URIs exactly for the file to open.
  final wpName =
      (String l) => XmlName(l, 'wp');
  final aName =
      (String l) => XmlName(l, 'a');
  final picName =
      (String l) => XmlName(l, 'pic');
  final rName =
      (String l) => XmlName(l, 'r');

  final extent = XmlElement(wpName('extent'), [
    XmlAttribute(XmlName('cx'), widthEmu.toString()),
    XmlAttribute(XmlName('cy'), heightEmu.toString()),
  ]);
  final docPr = XmlElement(wpName('docPr'), [
    XmlAttribute(XmlName('id'), docPrId.toString()),
    XmlAttribute(XmlName('name'), 'Picture $docPrId'),
  ]);
  final cNvGraphicFramePr = XmlElement(wpName('cNvGraphicFramePr'));

  final cNvPr = XmlElement(picName('cNvPr'), [
    XmlAttribute(XmlName('id'), '0'),
    XmlAttribute(XmlName('name'), 'placeholder.png'),
  ]);
  final cNvPicPr = XmlElement(picName('cNvPicPr'));
  final nvPicPr = XmlElement(picName('nvPicPr'), [], [cNvPr, cNvPicPr]);

  final blip = XmlElement(aName('blip'), [
    XmlAttribute(rName('embed'), relId),
  ]);
  final fillRect = XmlElement(aName('fillRect'));
  final stretch = XmlElement(aName('stretch'), [], [fillRect]);
  final blipFill = XmlElement(picName('blipFill'), [], [blip, stretch]);

  final off = XmlElement(aName('off'), [
    XmlAttribute(XmlName('x'), '0'),
    XmlAttribute(XmlName('y'), '0'),
  ]);
  final ext = XmlElement(aName('ext'), [
    XmlAttribute(XmlName('cx'), widthEmu.toString()),
    XmlAttribute(XmlName('cy'), heightEmu.toString()),
  ]);
  final xfrm = XmlElement(aName('xfrm'), [], [off, ext]);
  final avLst = XmlElement(aName('avLst'));
  final prstGeom = XmlElement(aName('prstGeom'), [
    XmlAttribute(XmlName('prst'), 'rect'),
  ], [avLst]);
  final spPr = XmlElement(picName('spPr'), [], [xfrm, prstGeom]);

  final pic = XmlElement(picName('pic'), [
    XmlAttribute(
      XmlName('pic', 'xmlns'),
      'http://schemas.openxmlformats.org/drawingml/2006/picture',
    ),
  ], [nvPicPr, blipFill, spPr]);

  final graphicData = XmlElement(aName('graphicData'), [
    XmlAttribute(
      XmlName('uri'),
      'http://schemas.openxmlformats.org/drawingml/2006/picture',
    ),
  ], [pic]);
  final graphic = XmlElement(aName('graphic'), [
    XmlAttribute(
      XmlName('a', 'xmlns'),
      'http://schemas.openxmlformats.org/drawingml/2006/main',
    ),
  ], [graphicData]);

  final inline = XmlElement(wpName('inline'), [
    XmlAttribute(XmlName('distT'), '0'),
    XmlAttribute(XmlName('distB'), '0'),
    XmlAttribute(XmlName('distL'), '0'),
    XmlAttribute(XmlName('distR'), '0'),
  ], [extent, docPr, cNvGraphicFramePr, graphic]);

  final drawing = XmlElement(w('drawing'), [], [inline]);
  final run = XmlElement(w('r'), [], [drawing]);
  final paragraph = XmlElement(w('p'), [], [run]);

  final sdtContent = XmlElement(w('sdtContent'), [], [paragraph]);

  return XmlElement(w('sdt'), [], [sdtPr, sdtContent]);
}
