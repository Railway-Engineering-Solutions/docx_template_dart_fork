library docx_template;

export 'src/template.dart'
    show DocxTemplate, DocxTemplateException, TagPolicy, ImagePolicy;
export 'src/model.dart';
export 'src/tag_models.dart' show TagType, DocxTag, DocxTagCollection;
export 'src/editor.dart'
    show
        DocxEditableStructure,
        DocxParagraphInfo,
        DocxTableInfo,
        TemplatedCell,
        TemplatedRow,
        SdtIdAllocator;
