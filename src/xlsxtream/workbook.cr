require "./errors"
require "./xml"
require "./shared_string_table"
require "./worksheet"
require "./zip_writer"

module Xlsxtream
  record FontOptions, name : String = "Calibri", size : Int32 = 12, family : String = "Swiss"

  class Workbook
    FONT_FAMILY_IDS = {
      ""           => 0,
      "roman"      => 1,
      "swiss"      => 2,
      "modern"     => 3,
      "script"     => 4,
      "decorative" => 5,
    }

    def self.open(output : String | Path | IO | ZipWriter,
                  use_shared_strings : Bool = false,
                  auto_format : Bool = false,
                  columns : Array(ColumnOptions)? = nil,
                  has_header_row : Bool = false,
                  font : FontOptions = FontOptions.new, & : Workbook ->) : Nil
      workbook = new(output, use_shared_strings: use_shared_strings,
        auto_format: auto_format, columns: columns,
        has_header_row: has_header_row, font: font)
      begin
        yield workbook
      ensure
        workbook.close
      end
    end

    def initialize(output : String | Path | IO | ZipWriter,
                   @use_shared_strings : Bool = false,
                   @auto_format : Bool = false,
                   @columns : Array(ColumnOptions)? = nil,
                   @has_header_row : Bool = false,
                   @font : FontOptions = FontOptions.new)
      @writer = ZipWriter.with_output_to(output)
      @sst = SharedStringTable.new
      @worksheets = [] of Worksheet
    end

    def add_worksheet(name : String? = nil,
                      use_shared_strings : Bool? = nil,
                      auto_format : Bool? = nil,
                      columns : Array(ColumnOptions)? = nil,
                      has_header_row : Bool? = nil) : Worksheet
      unless @worksheets.all? { |ws| ws.closed? }
        raise Error.new("Close the current worksheet before adding a new one")
      end

      build_worksheet(name,
        use_shared_strings: use_shared_strings,
        auto_format: auto_format,
        columns: columns,
        has_header_row: has_header_row)
    end

    def write_worksheet(name : String? = nil,
                        use_shared_strings : Bool? = nil,
                        auto_format : Bool? = nil,
                        columns : Array(ColumnOptions)? = nil,
                        has_header_row : Bool? = nil, & : Worksheet ->) : Nil
      worksheet = build_worksheet(name,
        use_shared_strings: use_shared_strings,
        auto_format: auto_format,
        columns: columns,
        has_header_row: has_header_row)
      yield worksheet
      worksheet.close
    end

    def write_worksheet(name : String? = nil,
                        use_shared_strings : Bool? = nil,
                        auto_format : Bool? = nil,
                        columns : Array(ColumnOptions)? = nil,
                        has_header_row : Bool? = nil) : Nil
      worksheet = build_worksheet(name,
        use_shared_strings: use_shared_strings,
        auto_format: auto_format,
        columns: columns,
        has_header_row: has_header_row)
      worksheet.close
    end

    def close : Nil
      write_workbook
      write_styles
      write_sst unless @sst.empty?
      write_workbook_rels
      write_root_rels
      write_content_types
      @writer.close
    end

    private def build_worksheet(name : String? = nil,
                                use_shared_strings : Bool? = nil,
                                auto_format : Bool? = nil,
                                columns : Array(ColumnOptions)? = nil,
                                has_header_row : Bool? = nil) : Worksheet
      use_sst = use_shared_strings.nil? ? @use_shared_strings : use_shared_strings
      af = auto_format.nil? ? @auto_format : auto_format
      cols = columns || @columns
      hhr = has_header_row.nil? ? @has_header_row : has_header_row
      sst = use_sst ? @sst : nil

      sheet_id = @worksheets.size + 1
      name = name || "Sheet#{sheet_id}"

      @writer.add_file "xl/worksheets/sheet#{sheet_id}.xml"

      worksheet = Worksheet.new(@writer, id: sheet_id, name: name, sst: sst, auto_format: af, columns: cols, has_header_row: hhr)
      @worksheets << worksheet

      worksheet
    end

    private def write_root_rels : Nil
      @writer.add_file "_rels/.rels"
      @writer << XML.header
      @writer << XML.strip(
        %(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">) +
        %(<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>) +
        %(</Relationships>)
      )
    end

    private def write_workbook : Nil
      rid = 0
      @writer.add_file "xl/workbook.xml"
      @writer << XML.header
      @writer << XML.strip(
        %(<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">) +
        %(<workbookPr date1904="false"/>) +
        %(<sheets>)
      )
      @worksheets.each do |worksheet|
        rid += 1
        @writer << %(<sheet name="#{XML.escape_attr worksheet.name.to_s}" sheetId="#{worksheet.id}" r:id="rId#{rid}"/>)
      end
      @writer << XML.strip(
        %(</sheets>) +
        %(</workbook>)
      )
    end

    private def write_styles : Nil
      font_size = @font.size.to_s
      font_name = @font.name
      font_family = @font.family.downcase
      font_family_id = FONT_FAMILY_IDS[font_family]? || raise Error.new(
        "Invalid font family #{font_family}, must be one of " +
        FONT_FAMILY_IDS.keys.map { |k| k.inspect }.join(", ")
      )

      @writer.add_file "xl/styles.xml"
      @writer << XML.header
      @writer << XML.strip(
        %(<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">) +
        %(<numFmts count="2">) +
        %(<numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>) +
        %(<numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>) +
        %(</numFmts>) +
        %(<fonts count="2">) +
        %(<font>) +
        %(<sz val="#{XML.escape_attr font_size}"/>) +
        %(<name val="#{XML.escape_attr font_name}"/>) +
        %(<family val="#{font_family_id}"/>) +
        %(</font>) +
        %(<font>) +
        %(<b val="1"/>) +
        %(<sz val="#{XML.escape_attr font_size}"/>) +
        %(<name val="#{XML.escape_attr font_name}"/>) +
        %(<family val="#{font_family_id}"/>) +
        %(</font>) +
        %(</fonts>) +
        %(<fills count="2">) +
        %(<fill>) +
        %(<patternFill patternType="none"/>) +
        %(</fill>) +
        %(<fill>) +
        %(<patternFill patternType="gray125"/>) +
        %(</fill>) +
        %(</fills>) +
        %(<borders count="1">) +
        %(<border/>) +
        %(</borders>) +
        %(<cellStyleXfs count="1">) +
        %(<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>) +
        %(</cellStyleXfs>) +
        %(<cellXfs count="6">) +
        %(<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>) +
        %(<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>) +
        %(<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>) +
        %(<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1">) +
        %(<alignment horizontal="center" vertical="center"/>) +
        %(</xf>) +
        %(<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">) +
        %(<alignment horizontal="center" vertical="center"/>) +
        %(</xf>) +
        %(<xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1">) +
        %(<alignment horizontal="center" vertical="center"/>) +
        %(</xf>) +
        %(</cellXfs>) +
        %(<cellStyles count="1">) +
        %(<cellStyle name="Normal" xfId="0" builtinId="0"/>) +
        %(</cellStyles>) +
        %(<dxfs count="0"/>) +
        %(<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>) +
        %(</styleSheet>)
      )
    end

    private def write_sst : Nil
      @writer.add_file "xl/sharedStrings.xml"
      @writer << XML.header
      @writer << %(<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{@sst.references}" uniqueCount="#{@sst.size}">)
      @sst.each_key do |string|
        @writer << "<si><t>#{XML.escape_value string}</t></si>"
      end
      @writer << "</sst>"
    end

    private def write_workbook_rels : Nil
      rid = 0
      @writer.add_file "xl/_rels/workbook.xml.rels"
      @writer << XML.header
      @writer << %(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)
      @worksheets.each do |worksheet|
        rid += 1
        @writer << %(<Relationship Id="rId#{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{worksheet.id}.xml"/>)
      end
      rid += 1
      @writer << %(<Relationship Id="rId#{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>)
      unless @sst.empty?
        rid += 1
        @writer << %(<Relationship Id="rId#{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>)
      end
      @writer << %(</Relationships>)
    end

    private def write_content_types : Nil
      @writer.add_file "[Content_Types].xml"
      @writer << XML.header
      @writer << XML.strip(
        %(<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">) +
        %(<Default Extension="xml" ContentType="application/xml"/>) +
        %(<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>) +
        %(<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>) +
        %(<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>)
      )
      unless @sst.empty?
        @writer << %(<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>)
      end
      @worksheets.each do |worksheet|
        @writer << %(<Override PartName="/xl/worksheets/sheet#{worksheet.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)
      end
      @writer << %(</Types>)
    end
  end
end
