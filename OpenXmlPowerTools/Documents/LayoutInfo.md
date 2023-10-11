
#### WordprocessingDocument
| App Path                                       | Xml Path                  |
| ---------------------------------------------- | ------------------------- |
| application/vnd.ms-word.styles.textEffects+xml | styles/style/@styleId     |
| application/vnd.ms-word.styles.textEffects+xml | styles/style/basedOn/@val |
| application/vnd.ms-word.styles.textEffects+xml | styles/style/link/@val    |
| application/vnd.ms-word.styles.textEffects+xml | styles/style/next/@val    |
| application/vnd.ms-word.stylesWithEffects+xml  | styles/style/@styleId     |
| application/vnd.ms-word.stylesWithEffects+xml  | styles/style/basedOn/@val |
| application/vnd.ms-word.stylesWithEffects+xml  | styles/style/link/@val    |
| application/vnd.ms-word.stylesWithEffects+xml  | styles/style/next/@val    |


| App Path                                                     | Xml Path            |
| ------------------------------------------------------------ | ------------------- |
| application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml | pPr/pStyle/@val     |
| application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml | rPr/rStyle/@val     |
| application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml | tblPr/tblStyle/@val |


| App Path                                                                   | Xml Path                  |
| -------------------------------------------------------------------------- | ------------------------- |
| application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml | pPr/pStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml | rPr/rStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml | tblPr/tblStyle/@val |


| App Path                                       | Xml Path                  |
| ---------------------------------------------- | ------------------------- |
|application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml| pPr/pStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml |                        rPr/rStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml |      tblPr/tblStyle/@val |


| App Path                                       | Xml Path                  |
| ---------------------------------------------- | ------------------------- |
|application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml| pPr/pStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml |        rPr/rStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml |       tblPr/tblStyle/@val |

| App Path                                       | Xml Path                  |
| ---------------------------------------------- | ------------------------- |
|application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml| pPr/pStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml |     rPr/rStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml |     tblPr/tblStyle/@val |

| App Path                                       | Xml Path                  |
| ---------------------------------------------- | ------------------------- |
|application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml| pPr/pStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml |       rPr/rStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml |       tblPr/tblStyle/@val |

| App Path                                       | Xml Path                  |
| ---------------------------------------------- | ------------------------- |
| application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml |    abstractNum/lvl/pStyle/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml |    abstractNum/numStyleLink/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml |    abstractNum/styleLink/@val |

| App Path                                                     | Xml Path                        |
| ------------------------------------------------------------ | ------------------------------- |
| application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml | settings/clickAndTypeStyle/@val |


| App Path                                       | Xml Path                  |
| ---------------------------------------------- | ------------------------- |
|application/vnd.ms-word.styles.textEffects+xml| styles/style/name/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.stylesWithEffects+xml                |styles/style/name/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml | styles/style/name/@val |
| application/vnd.ms-word.stylesWithEffects+xml |                                    styles/style/name/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.stylesWithEffects+xml | latentStyles/lsdException/@name |
| application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml | latentStyles/lsdException/@name |
| application/vnd.ms-word.stylesWithEffects+xml |                                    latentStyles/lsdException/@name |
| application/vnd.ms-word.styles.textEffects+xml |                                   latentStyles/lsdException/@name |

#### OpenXmlPart

| App Path                                                     | Xml Path                      |
| ------------------------------------------------------------ | ----------------------------- |
| application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml | abstractNum/lvl/pStyle/@val   |
| application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml | abstractNum/numStyleLink/@val |
| application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml | abstractNum/styleLink/@val    |

| App Path                                                     | Xml Path            |
| ------------------------------------------------------------ | ------------------- |
| application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml | pPr/pStyle/@val     |
| application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml | rPr/rStyle/@val     |
| application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml | tblPr/tblStyle/@val |

| App Path                                       | Xml Path                  |
| ---------------------------------------------- | ------------------------- |
| application/vnd.ms-word.styles.textEffects+xml | styles/style/@styleId     |
| application/vnd.ms-word.styles.textEffects+xml | styles/style/basedOn/@val |
| application/vnd.ms-word.styles.textEffects+xml | styles/style/link/@val    |
| application/vnd.ms-word.styles.textEffects+xml | styles/style/next/@val    |

#### WmlDocument
At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
querying or generating that markup.  This is an example of the GlossaryDocumentPart.

<w:glossaryDocument xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14">
  <w:docParts>
    <w:docPart>
      <w:docPartPr>
        <w:name w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
        <w:category>
          <w:name w:val="General" />
          <w:gallery w:val="placeholder" />
        </w:category>
        <w:types>
          <w:type w:val="bbPlcHdr" />
        </w:types>
        <w:behaviors>
          <w:behavior w:val="content" />
        </w:behaviors>
        <w:guid w:val="{13882A71-B5B7-4421-ACBB-9B61C61B3034}" />
      </w:docPartPr>
      <w:docPartBody>
        <w:p w:rsidR="00004EEA" w:rsidRDefault="00AD57F5" w:rsidP="00AD57F5">

At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
querying or generating that markup.  This is an example of the Custom XML Properties part.

<ds:datastoreItem ds:itemID="{1337A0C2-E6EE-4612-ACA5-E0E5A513381D}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
  <ds:schemaRefs />
</ds:datastoreItem>

At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
querying or generating that markup.  This is an example of the GlossaryDocument part.

<w:glossaryDocument xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14">
  <w:docParts>
    <w:docPart>
      <w:docPartPr>
        <w:name w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
        <w:category>
          <w:name w:val="General" />
          <w:gallery w:val="placeholder" />
        </w:category>
        <w:types>
          <w:type w:val="bbPlcHdr" />
        </w:types>
        <w:behaviors>
          <w:behavior w:val="content" />
        </w:behaviors>
        <w:guid w:val="{13882A71-B5B7-4421-ACBB-9B61C61B3034}" />
      </w:docPartPr>
      <w:docPartBody>
        <w:p w:rsidR="00004EEA" w:rsidRDefault="00AD57F5" w:rsidP="00AD57F5">
          <w:pPr>
            <w:pStyle w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
          </w:pPr>
          <w:r w:rsidRPr="00FB619D">
            <w:rPr>
              <w:rStyle w:val="PlaceholderText" />
              <w:lang w:val="da-DK" />
            </w:rPr>
            <w:t>Produktnavn</w:t>
          </w:r>
          <w:r w:rsidRPr="007379EE">
            <w:rPr>
              <w:rStyle w:val="PlaceholderText" />
            </w:rPr>
            <w:t>.</w:t>
          </w:r>
        </w:p>
      </w:docPartBody>
    </w:docPart>
  </w:docParts>
</w:glossaryDocument>


#### from the schema in the standard
writeProtection
view
zoom
removePersonalInformation
removeDateAndTime
doNotDisplayPageBoundaries
displayBackgroundShape
printPostScriptOverText
printFractionalCharacterWidth
printFormsData
embedTrueTypeFonts
embedSystemFonts
saveSubsetFonts
saveFormsData
mirrorMargins
alignBordersAndEdges
bordersDoNotSurroundHeader
bordersDoNotSurroundFooter
gutterAtTop
hideSpellingErrors
hideGrammaticalErrors
activeWritingStyle
proofState
formsDesign
attachedTemplate
linkStyles
stylePaneFormatFilter
stylePaneSortMethod
documentType
mailMerge
revisionView
trackRevisions
doNotTrackMoves
doNotTrackFormatting
documentProtection
autoFormatOverride
styleLockTheme
styleLockQFSet
defaultTabStop
autoHyphenation
consecutiveHyphenLimit
hyphenationZone
doNotHyphenateCaps
showEnvelope
summaryLength
clickAndTypeStyle
defaultTableStyle
evenAndOddHeaders
bookFoldRevPrinting
bookFoldPrinting
bookFoldPrintingSheets
drawingGridHorizontalSpacing
drawingGridVerticalSpacing
displayHorizontalDrawingGridEvery
displayVerticalDrawingGridEvery
doNotUseMarginsForDrawingGridOrigin
drawingGridHorizontalOrigin
drawingGridVerticalOrigin
doNotShadeFormData
noPunctuationKerning
characterSpacingControl
printTwoOnOne
strictFirstAndLastChars
noLineBreaksAfter
noLineBreaksBefore
savePreviewPicture
doNotValidateAgainstSchema
saveInvalidXml
ignoreMixedContent
alwaysShowPlaceholderText
doNotDemarcateInvalidXml
saveXmlDataOnly
useXSLTWhenSaving
saveThroughXslt
showXMLTags
alwaysMergeEmptyNamespace
updateFields
footnotePr
endnotePr
compat
docVars
rsids
m:mathPr
attachedSchema
themeFontLang
clrSchemeMapping
doNotIncludeSubdocsInStats
doNotAutoCompressPictures
forceUpgrade
captions
readModeInkLockDown
smartTagType
sl:schemaLibrary
doNotEmbedSmartTags
decimalSymbol
listSeparator