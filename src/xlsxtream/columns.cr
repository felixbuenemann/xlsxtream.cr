module Xlsxtream
  record ColumnOptions, width_chars : Int32? = nil, width_pixels : Float64? = nil

  class Columns
    def initialize(@columns : Array(ColumnOptions))
    end

    def to_xml : String
      xml = String::Builder.new
      xml << "<cols>"

      @columns.each_with_index do |column, index|
        width_chars = column.width_chars
        width_pixels = column.width_pixels

        if width_chars.nil? && width_pixels.nil?
          xml << %(<col min="#{index + 1}" max="#{index + 1}"/>)
        else
          width_pixels ||= ((((width_chars.not_nil! * 7.0) + 5) / 7) * 256).to_i / 256.0
          xml << %(<col min="#{index + 1}" max="#{index + 1}" width="#{width_pixels}" customWidth="1"/>)
        end
      end

      xml << "</cols>"
      xml.to_s
    end
  end
end
