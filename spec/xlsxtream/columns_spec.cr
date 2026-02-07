require "../spec_helper"

module Xlsxtream
  describe Columns do
    it "no width column" do
      column = Columns.new([ColumnOptions.new])
      expected = "<cols><col min=\"1\" max=\"1\"/></cols>"
      column.to_xml.should eq(expected)
    end

    it "pixel width column" do
      column = Columns.new([ColumnOptions.new(width_pixels: 2341.5)])
      expected = "<cols><col min=\"1\" max=\"1\" width=\"2341.5\" customWidth=\"1\"/></cols>"
      column.to_xml.should eq(expected)
    end

    it "character width column" do
      # https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.column.aspx
      #
      # ...Therefore, if the cell width is 8 characters wide, the value of
      # this attribute must be Truncate([8*7+5]/7*256)/256 = 8.7109375...
      column = Columns.new([ColumnOptions.new(width_chars: 8)])
      expected = "<cols><col min=\"1\" max=\"1\" width=\"8.7109375\" customWidth=\"1\"/></cols>"
      column.to_xml.should eq(expected)
    end

    it "mixed columns" do
      column = Columns.new([
        ColumnOptions.new,
        ColumnOptions.new(width_pixels: 61.0),
        ColumnOptions.new(width_chars: 14),
      ])
      expected = "<cols><col min=\"1\" max=\"1\"/><col min=\"2\" max=\"2\" width=\"61.0\" customWidth=\"1\"/><col min=\"3\" max=\"3\" width=\"14.7109375\" customWidth=\"1\"/></cols>"
      column.to_xml.should eq(expected)
    end
  end
end
