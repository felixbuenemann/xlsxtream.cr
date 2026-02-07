require "./xml"
require "./row"

module Xlsxtream
  class Worksheet
    getter id : Int32?
    getter name : String?

    def initialize(@io : IO | ZipWriter,
                   id : Int32? = nil,
                   name : String? = nil,
                   sst : SharedStringTable? = nil,
                   auto_format : Bool = false,
                   columns : Array(ColumnOptions)? = nil,
                   has_header_row : Bool = false)
      @id = id
      @name = name
      @sst = sst
      @auto_format = auto_format
      @columns = columns
      @has_header_row = has_header_row
      @rownum = 1
      @closed = false

      write_header
    end

    def <<(row : Array(CellValue)) : Nil
      is_header = @has_header_row && @rownum == 1
      @io << Row.new(row, @rownum, sst: @sst, auto_format: @auto_format, is_header: is_header).to_xml
      @rownum += 1
    end

    def add_row(row : Array(CellValue)) : Nil
      self << row
    end

    def close : Nil
      write_footer
      @closed = true
    end

    def closed? : Bool
      @closed
    end

    private def write_header : Nil
      @io << XML.header
      @io << XML.strip(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      )

      if columns = @columns
        unless columns.empty?
          @io << Columns.new(columns).to_xml
        end
      end

      @io << XML.strip(
        "<sheetData>"
      )
    end

    private def write_footer : Nil
      @io << XML.strip(
        "</sheetData></worksheet>"
      )
    end
  end
end
