require "../spec_helper"

module Xlsxtream
  describe Row do
    # =========================================================================
    # Normal rows
    # =========================================================================

    it "empty column" do
      row = Row.new([nil] of CellValue, 1)
      expected = "<row r=\"1\"></row>"
      row.to_xml.should eq(expected)
    end

    it "string column" do
      row = Row.new(["hello"] of CellValue, 1)
      expected = "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>hello</t></is></c></row>"
      row.to_xml.should eq(expected)
    end

    it "boolean column" do
      row = Row.new([true] of CellValue, 1)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"b\"><v>1</v></c></row>")
      row = Row.new([false] of CellValue, 1)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"b\"><v>0</v></c></row>")
    end

    it "text boolean column" do
      row = Row.new(["true"] of CellValue, 1, auto_format: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"b\"><v>1</v></c></row>")
      row = Row.new(["false"] of CellValue, 1, auto_format: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"b\"><v>0</v></c></row>")
    end

    it "integer column" do
      row = Row.new([1] of CellValue, 1)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"n\"><v>1</v></c></row>")
    end

    it "text integer column" do
      row = Row.new(["1"] of CellValue, 1, auto_format: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"n\"><v>1</v></c></row>")
    end

    it "float column" do
      row = Row.new([1.5] of CellValue, 1)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"n\"><v>1.5</v></c></row>")
    end

    it "text float column" do
      row = Row.new(["1.5"] of CellValue, 1, auto_format: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"n\"><v>1.5</v></c></row>")
    end

    it "date column" do
      row = Row.new([Xlsxtream::Date.new(1900, 1, 1)] of CellValue, 1)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"1\"><v>2.0</v></c></row>")
    end

    it "text date column" do
      row = Row.new(["1900-01-01"] of CellValue, 1, auto_format: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"1\"><v>2.0</v></c></row>")
    end

    it "invalid text date column" do
      row = Row.new(["1900-02-29"] of CellValue, 1, auto_format: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>1900-02-29</t></is></c></row>")
    end

    it "date time column" do
      row = Row.new([Time.utc(1900, 1, 1, 12, 0, 0)] of CellValue, 1)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"2\"><v>2.5</v></c></row>")
    end

    it "text date time column" do
      candidates = [
        "1900-01-01T12:00",
        "1900-01-01T12:00Z",
        "1900-01-01T12:00+00:00",
        "1900-01-01T12:00:00+00:00",
        "1900-01-01T12:00:00.000+00:00",
        "1900-01-01T12:00:00.000000000Z",
      ]
      candidates.each do |timestamp|
        row = Row.new([timestamp] of CellValue, 1, auto_format: true)
        row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"2\"><v>2.5</v></c></row>")
      end
      row = Row.new(["1900-01-01T12"] of CellValue, 1, auto_format: true)
      row.to_xml.should_not eq("<row r=\"1\"><c r=\"A1\" s=\"2\"><v>2.5</v></c></row>")
    end

    it "invalid text date time column" do
      row = Row.new(["1900-02-29T12:00"] of CellValue, 1, auto_format: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>1900-02-29T12:00</t></is></c></row>")
    end

    it "time column" do
      row = Row.new([Time.utc(1900, 1, 1, 12, 0, 0)] of CellValue, 1)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"2\"><v>2.5</v></c></row>")
    end

    it "string column with shared string table" do
      sst = SharedStringTable.new
      sst["hello"]
      row = Row.new(["hello"] of CellValue, 1, sst: sst)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>")
    end

    it "multiple columns" do
      row = Row.new(["foo", nil, 23] of CellValue, 1)
      expected = "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>foo</t></is></c><c r=\"C1\" t=\"n\"><v>23</v></c></row>"
      row.to_xml.should eq(expected)
    end

    # =========================================================================
    # Header rows
    # =========================================================================

    it "header string column" do
      row = Row.new(["hello"] of CellValue, 1, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"inlineStr\"><is><t>hello</t></is></c></row>")
    end

    it "header boolean column" do
      row = Row.new([true] of CellValue, 1, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"b\"><v>1</v></c></row>")
      row = Row.new([false] of CellValue, 1, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"b\"><v>0</v></c></row>")
    end

    it "header text boolean column" do
      row = Row.new(["true"] of CellValue, 1, auto_format: true, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"b\"><v>1</v></c></row>")
      row = Row.new(["false"] of CellValue, 1, auto_format: true, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"b\"><v>0</v></c></row>")
    end

    it "header integer column" do
      row = Row.new([1] of CellValue, 1, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"n\"><v>1</v></c></row>")
    end

    it "header text integer column" do
      row = Row.new(["1"] of CellValue, 1, auto_format: true, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"n\"><v>1</v></c></row>")
    end

    it "header float column" do
      row = Row.new([1.5] of CellValue, 1, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"n\"><v>1.5</v></c></row>")
    end

    it "header text float column" do
      row = Row.new(["1.5"] of CellValue, 1, auto_format: true, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"n\"><v>1.5</v></c></row>")
    end

    it "header date column" do
      row = Row.new([Xlsxtream::Date.new(1900, 1, 1)] of CellValue, 1, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"4\"><v>2.0</v></c></row>")
    end

    it "header text date column" do
      row = Row.new(["1900-01-01"] of CellValue, 1, auto_format: true, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"4\"><v>2.0</v></c></row>")
    end

    it "header invalid text date column" do
      row = Row.new(["1900-02-29"] of CellValue, 1, auto_format: true, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"inlineStr\"><is><t>1900-02-29</t></is></c></row>")
    end

    it "header date time column" do
      row = Row.new([Time.utc(1900, 1, 1, 12, 0, 0)] of CellValue, 1, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"5\"><v>2.5</v></c></row>")
    end

    it "header text date time column" do
      candidates = [
        "1900-01-01T12:00",
        "1900-01-01T12:00Z",
        "1900-01-01T12:00+00:00",
        "1900-01-01T12:00:00+00:00",
        "1900-01-01T12:00:00.000+00:00",
        "1900-01-01T12:00:00.000000000Z",
      ]
      candidates.each do |timestamp|
        row = Row.new([timestamp] of CellValue, 1, auto_format: true, is_header: true)
        row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"5\"><v>2.5</v></c></row>")
      end
      row = Row.new(["1900-01-01T12"] of CellValue, 1, auto_format: true, is_header: true)
      row.to_xml.should_not eq("<row r=\"1\"><c r=\"A1\" s=\"5\"><v>2.5</v></c></row>")
    end

    it "header invalid text date time column" do
      row = Row.new(["1900-02-29T12:00"] of CellValue, 1, auto_format: true, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"inlineStr\"><is><t>1900-02-29T12:00</t></is></c></row>")
    end

    it "header time column" do
      row = Row.new([Time.utc(1900, 1, 1, 12, 0, 0)] of CellValue, 1, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"5\"><v>2.5</v></c></row>")
    end

    it "header string column with shared string table" do
      sst = SharedStringTable.new
      sst["hello"]
      row = Row.new(["hello"] of CellValue, 1, sst: sst, is_header: true)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c></row>")
    end

    it "header multiple columns" do
      row = Row.new(["foo", nil, 23] of CellValue, 1, is_header: true)
      expected = "<row r=\"1\"><c r=\"A1\" s=\"3\" t=\"inlineStr\"><is><t>foo</t></is></c><c r=\"C1\" s=\"3\" t=\"n\"><v>23</v></c></row>"
      row.to_xml.should eq(expected)
    end
  end
end
