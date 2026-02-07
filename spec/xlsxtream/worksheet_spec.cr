require "../spec_helper"

module Xlsxtream
  describe Worksheet do
    it "empty worksheet" do
      io = IO::Memory.new
      ws = Worksheet.new(io)
      ws.close
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData></sheetData></worksheet>"
      io.to_s.should eq(expected)
    end

    it "add row" do
      io = IO::Memory.new
      ws = Worksheet.new(io)
      ws << ["foo"] of CellValue
      ws.add_row ["bar"] of CellValue
      ws.close
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" \
        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>foo</t></is></c></row>" \
        "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>bar</t></is></c></row>" \
        "</sheetData></worksheet>"
      io.to_s.should eq(expected)
    end

    it "add row with sst option" do
      io = IO::Memory.new
      sst = SharedStringTable.new
      sst["foo"]
      ws = Worksheet.new(io, sst: sst)
      ws << ["foo"] of CellValue
      ws.close
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" \
        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>" \
        "</sheetData></worksheet>"
      io.to_s.should eq(expected)
    end

    it "add row with auto_format option" do
      io = IO::Memory.new
      ws = Worksheet.new(io, auto_format: true)
      ws << ["1.5"] of CellValue
      ws.close
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" \
        "<row r=\"1\"><c r=\"A1\" t=\"n\"><v>1.5</v></c></row>" \
        "</sheetData></worksheet>"
      io.to_s.should eq(expected)
    end

    it "add row with has_header_row option" do
      io = IO::Memory.new
      ws = Worksheet.new(io, has_header_row: true)
      ws << ["header"] of CellValue
      ws.add_row ["not header"] of CellValue
      ws.close
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" \
        "<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"inlineStr\"><is><t>header</t></is></c></row>" \
        "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>not header</t></is></c></row>" \
        "</sheetData></worksheet>"
      io.to_s.should eq(expected)
    end

    it "add columns via worksheet options" do
      io = IO::Memory.new
      ws = Worksheet.new(io, columns: [
        ColumnOptions.new,
        ColumnOptions.new,
        ColumnOptions.new(width_pixels: 42.0),
      ])
      ws.close
      expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" \
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><cols>" \
        "<col min=\"1\" max=\"1\"/>" \
        "<col min=\"2\" max=\"2\"/>" \
        "<col min=\"3\" max=\"3\" width=\"42.0\" customWidth=\"1\"/>" \
        "</cols>" \
        "<sheetData></sheetData></worksheet>"
      io.to_s.should eq(expected)
    end

    it "add columns via worksheet options and add rows" do
      io = IO::Memory.new
      ws = Worksheet.new(io, columns: [
        ColumnOptions.new,
        ColumnOptions.new,
        ColumnOptions.new(width_pixels: 42.0),
      ])
      ws << ["foo"] of CellValue
      ws.add_row ["bar"] of CellValue
      ws.close
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
      io.to_s.should eq(expected)
    end

    it "respond to id" do
      ws = Worksheet.new(IO::Memory.new, id: 1)
      ws.id.should eq(1)
    end

    it "respond to name" do
      ws = Worksheet.new(IO::Memory.new, name: "test")
      ws.name.should eq("test")
    end
  end
end
