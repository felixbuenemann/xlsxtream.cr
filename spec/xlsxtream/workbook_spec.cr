require "../spec_helper"

module Xlsxtream
  class SpyWriter < Xlsxtream::ZipWriter
    @paths_to_file_contents = Hash(String, String).new
    @current = ""

    def initialize
      @zip = uninitialized Compress::Zip::Writer
      @close = [] of Closeable
      @current_writer = nil
      @entry_done = nil
      @buffer = String::Builder.new
    end

    def <<(data : String) : self
      @paths_to_file_contents[@current] = (@paths_to_file_contents[@current]? || "") + data
      self
    end

    def add_file(path : String) : Nil
      @current = path
      @paths_to_file_contents[@current] = ""
    end

    def [](key : String) : String
      @paths_to_file_contents[key]
    end

    def close : Nil
    end
  end

  describe Workbook do
    it "workbook from path" do
      tempfile = File.tempfile("xlsxtream")
      begin
        Workbook.open(tempfile.path) { }
        File.size(tempfile.path).should_not eq(0)
      ensure
        tempfile.delete
      end
    end

    it "workbook from io" do
      tempfile = File.tempfile("xlsxtream")
      begin
        io = File.open(tempfile.path, "wb")
        Workbook.open(io) { }
        io.close
        File.size(tempfile.path).should_not eq(0)
      ensure
        tempfile.delete
      end
    end

    it "empty workbook" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) { }
      expected_workbook =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " \
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" \
        "<workbookPr date1904=\"false\"/>" \
        "<sheets></sheets>" \
        "</workbook>"
      expected_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" \
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" \
        "</Relationships>"
      iow_spy["xl/workbook.xml"].should eq(expected_workbook)
      iow_spy["xl/_rels/workbook.xml.rels"].should eq(expected_rels)
    end

    it "workbook with sheet" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet
      end
      expected_sheet =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" \
        "<sheetData></sheetData>" \
        "</worksheet>"
      expected_workbook =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " \
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" \
        "<workbookPr date1904=\"false\"/>" \
        "<sheets>" \
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>" \
        "</sheets>" \
        "</workbook>"
      expected_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" \
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" \
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" \
        "</Relationships>"
      iow_spy["xl/worksheets/sheet1.xml"].should eq(expected_sheet)
      iow_spy["xl/workbook.xml"].should eq(expected_workbook)
      iow_spy["xl/_rels/workbook.xml.rels"].should eq(expected_rels)
    end

    it "workbook with sheet without block" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) do |wb|
        ws = wb.add_worksheet
        ws << ["foo"] of CellValue
        ws.close
      end
      expected_sheet =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" \
        "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>foo</t></is></c></row></sheetData>" \
        "</worksheet>"
      expected_workbook =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " \
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" \
        "<workbookPr date1904=\"false\"/>" \
        "<sheets>" \
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>" \
        "</sheets>" \
        "</workbook>"
      expected_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" \
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" \
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" \
        "</Relationships>"
      iow_spy["xl/worksheets/sheet1.xml"].should eq(expected_sheet)
      iow_spy["xl/workbook.xml"].should eq(expected_workbook)
      iow_spy["xl/_rels/workbook.xml.rels"].should eq(expected_rels)
    end

    it "workbook with sst" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet(use_shared_strings: true) do |ws|
          ws << ["foo"] of CellValue
        end
      end
      expected_sheet =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" \
        "<sheetData>" \
        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>" \
        "</sheetData>" \
        "</worksheet>"
      expected_workbook =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " \
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" \
        "<workbookPr date1904=\"false\"/>" \
        "<sheets>" \
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>" \
        "</sheets>" \
        "</workbook>"
      expected_sst =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">" \
        "<si><t>foo</t></si>" \
        "</sst>"
      expected_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" \
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" \
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" \
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" \
        "</Relationships>"
      iow_spy["xl/worksheets/sheet1.xml"].should eq(expected_sheet)
      iow_spy["xl/workbook.xml"].should eq(expected_workbook)
      iow_spy["xl/sharedStrings.xml"].should eq(expected_sst)
      iow_spy["xl/_rels/workbook.xml.rels"].should eq(expected_rels)
    end

    it "root relations" do
      iow_spy = SpyWriter.new
      wb = Workbook.new(iow_spy)
      wb.close
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" \
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" \
        "</Relationships>"
      iow_spy["_rels/.rels"].should eq(expected)
    end

    it "content types" do
      iow_spy = SpyWriter.new
      wb = Workbook.new(iow_spy)
      wb.close
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" \
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" \
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" \
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" \
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" \
        "</Types>"
      iow_spy["[Content_Types].xml"].should eq(expected)
    end

    it "write multiple worksheets" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet
        wb.write_worksheet
      end

      expected_workbook =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " \
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" \
        "<workbookPr date1904=\"false\"/>" \
        "<sheets>" \
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>" \
        "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>" \
        "</sheets>" \
        "</workbook>"
      expected_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" \
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" \
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>" \
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" \
        "</Relationships>"
      expected_sheet1 =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData></sheetData></worksheet>"
      expected_sheet2 =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData></sheetData></worksheet>"
      iow_spy["xl/workbook.xml"].should eq(expected_workbook)
      iow_spy["xl/_rels/workbook.xml.rels"].should eq(expected_rels)
      iow_spy["xl/worksheets/sheet1.xml"].should eq(expected_sheet1)
      iow_spy["xl/worksheets/sheet2.xml"].should eq(expected_sheet2)
    end

    it "must write sequentially" do
      iow_spy1 = SpyWriter.new
      Workbook.open(iow_spy1) do |wb|
        ws = wb.add_worksheet
        ws.close
        ws = wb.add_worksheet
        ws.close
      end

      iow_spy2 = SpyWriter.new
      expect_raises(Xlsxtream::Error) do
        Workbook.open(iow_spy2) do |wb|
          wb.add_worksheet
          wb.add_worksheet # adding a second worksheet without closing
        end
      end
    end

    it "write named worksheet" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet("foo")
      end

      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " \
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" \
        "<workbookPr date1904=\"false\"/>" \
        "<sheets>" \
        "<sheet name=\"foo\" sheetId=\"1\" r:id=\"rId1\"/>" \
        "</sheets>" \
        "</workbook>"
      iow_spy["xl/workbook.xml"].should eq(expected)
    end

    it "write unnamed worksheet with options" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet(use_shared_strings: true)
      end

      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " \
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" \
        "<workbookPr date1904=\"false\"/>" \
        "<sheets>" \
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>" \
        "</sheets>" \
        "</workbook>"
      iow_spy["xl/workbook.xml"].should eq(expected)
    end

    it "worksheet name as option" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) do |wb|
        wb.write_worksheet(name: "foo")
      end
      expected = "<sheet name=\"foo\" sheetId=\"1\" r:id=\"rId1\"/>"
      actual = iow_spy["xl/workbook.xml"]
      match = actual.match(/<sheet [^>]+>/)
      match.should_not be_nil
      match.not_nil![0].should eq(expected)
    end

    it "add columns via workbook options" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy, columns: [
        ColumnOptions.new,
        ColumnOptions.new,
        ColumnOptions.new(width_pixels: 42.0),
      ]) do |wb|
        wb.write_worksheet { }
      end

      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cols>" \
        "<col min=\"1\" max=\"1\"/>" \
        "<col min=\"2\" max=\"2\"/>" \
        "<col min=\"3\" max=\"3\" width=\"42.0\" customWidth=\"1\"/>" \
        "</cols>" \
        "<sheetData></sheetData></worksheet>"

      iow_spy["xl/worksheets/sheet1.xml"].should eq(expected)
    end

    it "add columns via workbook options and add rows" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy, columns: [
        ColumnOptions.new,
        ColumnOptions.new,
        ColumnOptions.new(width_pixels: 42.0),
      ]) do |wb|
        wb.write_worksheet do |ws|
          ws << ["foo"] of CellValue
          ws.add_row ["bar"] of CellValue
        end
      end

      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cols>" \
        "<col min=\"1\" max=\"1\"/>" \
        "<col min=\"2\" max=\"2\"/>" \
        "<col min=\"3\" max=\"3\" width=\"42.0\" customWidth=\"1\"/>" \
        "</cols>" \
        "<sheetData>" \
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>foo</t></is></c></row>" \
        "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>bar</t></is></c></row>" \
        "</sheetData></worksheet>"

      iow_spy["xl/worksheets/sheet1.xml"].should eq(expected)
    end

    it "styles content" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy) { }
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" \
        "<numFmts count=\"2\">" \
        "<numFmt numFmtId=\"164\" formatCode=\"yyyy\\-mm\\-dd\"/>" \
        "<numFmt numFmtId=\"165\" formatCode=\"yyyy\\-mm\\-dd hh:mm:ss\"/>" \
        "</numFmts>" \
        "<fonts count=\"2\">" \
        "<font>" \
        "<sz val=\"12\"/>" \
        "<name val=\"Calibri\"/>" \
        "<family val=\"2\"/>" \
        "</font>" \
        "<font>" \
        "<b val=\"1\"/>" \
        "<sz val=\"12\"/>" \
        "<name val=\"Calibri\"/>" \
        "<family val=\"2\"/>" \
        "</font>" \
        "</fonts>" \
        "<fills count=\"2\">" \
        "<fill>" \
        "<patternFill patternType=\"none\"/>" \
        "</fill>" \
        "<fill>" \
        "<patternFill patternType=\"gray125\"/>" \
        "</fill>" \
        "</fills>" \
        "<borders count=\"1\">" \
        "<border/>" \
        "</borders>" \
        "<cellStyleXfs count=\"1\">" \
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" \
        "</cellStyleXfs>" \
        "<cellXfs count=\"6\">" \
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" \
        "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" \
        "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" \
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\">" \
        "<alignment horizontal=\"center\" vertical=\"center\"/>" \
        "</xf>" \
        "<xf numFmtId=\"164\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyNumberFormat=\"1\">" \
        "<alignment horizontal=\"center\" vertical=\"center\"/>" \
        "</xf>" \
        "<xf numFmtId=\"165\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\" applyNumberFormat=\"1\">" \
        "<alignment horizontal=\"center\" vertical=\"center\"/>" \
        "</xf>" \
        "</cellXfs>" \
        "<cellStyles count=\"1\">" \
        "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>" \
        "</cellStyles>" \
        "<dxfs count=\"0\"/>" \
        "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>" \
        "</styleSheet>"
      iow_spy["xl/styles.xml"].should eq(expected)
    end

    it "custom font size" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy, font: FontOptions.new(size: 23)) { }
      expected = "<sz val=\"23\"/>"
      match = iow_spy["xl/styles.xml"].match(/<sz [^>]+>/)
      match.should_not be_nil
      match.not_nil![0].should eq(expected)
    end

    it "custom font name" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy, font: FontOptions.new(name: "Comic Sans")) { }
      expected = "<name val=\"Comic Sans\"/>"
      match = iow_spy["xl/styles.xml"].match(/<name [^>]+>/)
      match.should_not be_nil
      match.not_nil![0].should eq(expected)
    end

    it "custom font family" do
      iow_spy = SpyWriter.new
      Workbook.open(iow_spy, font: FontOptions.new(family: "Script")) { }
      expected = "<family val=\"4\"/>"
      match = iow_spy["xl/styles.xml"].match(/<family [^>]+>/)
      match.should_not be_nil
      match.not_nil![0].should eq(expected)
    end

    it "font family mapping" do
      tests = {
        ""           => 0,
        "ROMAN"      => 1,
        "Roman"      => 1,
        "swiss"      => 2,
        "modern"     => 3,
        "script"     => 4,
        "decorative" => 5,
      }
      tests.each do |value, id|
        iow_spy = SpyWriter.new
        Workbook.open(iow_spy, font: FontOptions.new(family: value)) { }
        expected = "<family val=\"#{id}\"/>"
        match = iow_spy["xl/styles.xml"].match(/<family [^>]+>/)
        match.should_not be_nil
        match.not_nil![0].should eq(expected)
      end
    end

    it "invalid font family" do
      iow_spy = SpyWriter.new
      expect_raises(Xlsxtream::Error) do
        Workbook.open(iow_spy, font: FontOptions.new(family: "Foo")) { }
      end
    end
  end
end
