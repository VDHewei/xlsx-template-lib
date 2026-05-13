/** Office 文档关系的命名空间 */
export const DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/** 计算链关系的命名空间 */
export const CALC_CHAIN_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

/** 共享字符串关系的命名空间 */
export const SHARED_STRINGS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

/** 超链接关系的命名空间 */
export const HYPERLINK_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

// ==================== RichData（富数据/本地图片）相关常量 ====================

/** RichData 关系文件的 XML 模板 */
export const RICH_DATA_xml_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;

/** RichData 值文件的 XML 模板 */
export const RICH_DATA_xml_RV = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="0">
</rvData>`;

/** RichData 结构文件的 XML 模板（含 _localImage 类型定义） */
export const RICH_DATA_xml_STRUCTURE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
    <s t="_localImage">
        <k n="_rvRel:LocalImageIdentifier" t="i"/>
        <k n="CalcOrigin" t="i"/>
    </s>
</rvStructures>`;

/** RichData 类型信息文件的 XML 模板 */
export const RICH_DATA_xml_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x"
    xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <global>
        <keyFlags>
            <key name="_Self">
                <flag name="ExcludeFromFile" value="1"/>
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_DisplayString">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_Flags">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_Format">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_SubLabel">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_Attribution">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_Icon">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_Display">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_CanonicalPropertyNames">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
            <key name="_ClassificationId">
                <flag name="ExcludeFromCalcComparison" value="1"/>
            </key>
        </keyFlags>
    </global>
</rvTypesInfo>`;

/** RichData 值关系文件的 XML 模板 */
export const RICH_DATA_xml_VALUE_REL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
</richValueRels>`;

/** RichData 元数据文件的 XML 模板 */
export const RICH_DATA_xml_METADATA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
    <metadataTypes count="1">
        <metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>
    </metadataTypes>
    <futureMetadata name="XLRICHVALUE" count="0">
    </futureMetadata>
    <valueMetadata count="0">
    </valueMetadata>
</metadata>`;

/** RichData 关系文件路径（archive 内） */
export const RICH_DATA_RELS_FILE = 'xl/richData/_rels/richValueRel.xml.rels';
/** RichData 值文件路径 */
export const RICH_DATA_RV_FILE = 'xl/richData/rdrichvalue.xml';
/** RichData 结构文件路径 */
export const RICH_DATA_STRUCTURE_FILE = 'xl/richData/rdrichvaluestructure.xml';
/** RichData 类型信息文件路径 */
export const RICH_DATA_TYPES_FILE = 'xl/richData/rdRichValueTypes.xml';
/** RichData 值关系文件路径 */
export const RICH_DATA_VALUE_REL_FILE = 'xl/richData/richValueRel.xml';
/** RichData 元数据文件路径 */
export const RICH_DATA_METADATA_FILE = 'xl/metadata.xml';
