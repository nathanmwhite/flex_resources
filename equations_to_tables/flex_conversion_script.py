# TODO: add initial column in each row such that content is seven rows wide, and example number is inserted

import os
import re

def get_content(file_data):
    found = []
    example = []

    base_word = ""
    gloss = ""
    for item in file_data:
        if "Interlin Phrase Number" in item:
            phrase_number_searched = re.search('<w:t>([0-9\.]+?)</w:t>', item)
            if phrase_number_searched:
                example.append(("("+str(phrase_number_searched.group(1))+")",''))
#                print("number appended")
        elif "Interlin Base" in item:
            base_word_searched = re.search("<m:t>(.*?)</m:t>", item)
            if base_word_searched:
                base_word = base_word_searched.group(1)[:-2]
        elif "Interlin Word Gloss" in item:
            gloss_searched = re.search("<m:t>(.*?)</m:t>", item)
            if gloss_searched:
                gloss = gloss_searched.group(1)[:-2]
        elif "Interlin Freeform" in item:
            freeform_searched = re.search("<w:t>(.*?)</w:t>", item)
            if freeform_searched:
                freeform = freeform_searched.group(1)
            example.append((freeform,))
            found.append(example)
            example = []
        elif '<w:t xml:space="preserve">' in item:
            if base_word in [',', '.', '..', '...'] and len(example) > 1:
                example[-1] = (example[-1][0]+base_word, example[-1][1])
            else:
                example.append((base_word, gloss))
            base_word = ""
            gloss = ""
    return found

header = """<?xml version="1.0" encoding="utf-8"?>
<?mso-application  progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" xmlns="http://www.w3.org/1999/XSL/Format" xmlns:v="urn:schemas-microsoft-com:vml">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="256">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
        <w:body>
"""

footer = """
</w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
"""

gridCol = """<w:gridCol w:w="1127"/>"""

gridSpan = """<w:gridSpan w:val="{num_cols}"/>"""

table_header = """
<w:tbl>
   <w:tblPr>
      <w:tblStyle w:val="TableGrid"/>
      <w:tblW w:w="0" w:type="auto"/>
      <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
   </w:tblPr>
   <w:tblGrid>
      {gridCol}
   </w:tblGrid>
"""

table_row_start = """   <w:tr w:rsidR="00E943E8" w:rsidTr="00E943E8">"""

table_cell = """<w:tc>
         <w:tcPr>
            <w:tcW w:w="{cell_width}" w:type="dxa"/>
            {grid_span}
         </w:tcPr>
         <w:p w:rsidR="00E943E8" w:rsidRDefault="00E943E8">
            <w:r>
               <w:t>{content}</w:t>
            </w:r>
         </w:p>
      </w:tc>
"""

table_row_end = """   </w:tr>"""

table_footer = """</w:tbl>"""

p = """<w:p w:rsidR="00E943E8" w:rsidRDefault="00E943E8"><w:r><w:t>{content}</w:t></w:r></w:p>"""

CELL_WIDTH = 1127

def populate_content(processed_data):
    line_length = 8
    document = ""
    document += header
    for item in processed_data:
        total_full_lines = max(0, (len(item) - 1) // line_length)
        over_line_length = max(0, (len(item) - 1) % line_length)
        rows = list(zip(*item[:-1]))
        for i in range(total_full_lines):
            if i == 0:
                document += table_header.format(gridCol=gridCol * line_length)
            else:
                document += table_header.format(gridCol=gridCol * (line_length + 1))
            document += table_row_start
            if i != 0:
                document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content="")
            for word in rows[0][i*line_length:(i+1)*line_length]:
                document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content=word)
            document += table_row_end
            document += table_row_start
            if i != 0:
                document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content="")
            for word in rows[1][i*line_length:(i+1)*line_length]:
                document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content=word)
            document += table_row_end
            if over_line_length == 0 and i == total_full_lines - 1:
                document += table_row_start
                document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content="")
                if i == 0:
                    document += table_cell.format(cell_width=CELL_WIDTH * (line_length - 1), grid_span=gridSpan.format(num_cols=line_length-1), content=item[-1][0])
                else:
                    document += table_cell.format(cell_width=CELL_WIDTH * (line_length), grid_span=gridSpan.format(num_cols=line_length), content=item[-1][0])
                document += table_row_end
            document += table_footer
            document += p.format(content="")
        if over_line_length > 0:
            single_line = (len(item) < line_length)
            document += table_header.format(gridCol=gridCol * line_length)
            document += table_row_start
            if not single_line:
                document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content="")
            for i in range(over_line_length):
                if i == over_line_length - 1:
                    document += table_cell.format(cell_width=CELL_WIDTH * (line_length - i), grid_span=gridSpan.format(num_cols=line_length-i), content=rows[0][-1])
                else:
                    document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content=rows[0][total_full_lines*line_length+i])
            document += table_row_end
            document += table_row_start
            if not single_line:
                document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content="")
            for i in range(over_line_length):
                if i == over_line_length - 1:
                    document += table_cell.format(cell_width=CELL_WIDTH * (line_length - i), grid_span=gridSpan.format(num_cols=line_length-i), content=rows[1][-1])
                else:
                    document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content=rows[1][total_full_lines*line_length+i])
            document += table_row_end
            document += table_row_start
            document += table_cell.format(cell_width=CELL_WIDTH, grid_span="", content="")
            if not single_line:
                document += table_cell.format(cell_width=CELL_WIDTH * (line_length - 1), grid_span=gridSpan.format(num_cols=line_length-1), content=item[-1][0])
            else:
                document += table_cell.format(cell_width=CELL_WIDTH * (line_length), grid_span=gridSpan.format(num_cols=line_length), content=item[-1][0])
            document += table_row_end
            document += table_footer
            document += p.format(content="")
    document += footer
    return document

if __name__ == '__main__':
    os.chdir(os.path.expanduser(os.path.join('~', 'Documents')))
    f = open('Interlinear texts.xml', 'r', encoding="iso-8859-1")
    data = get_content(f.readlines())
    f.close()
    content = populate_content(data)
    f = open('test_document.xml', 'w', encoding="iso-8859-1")
    f.write(''.join(content))
    f.close()
