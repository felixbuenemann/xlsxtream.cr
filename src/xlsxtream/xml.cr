module Xlsxtream
  module XML
    XML_ESCAPES = {
      '&' => "&amp;",
      '"' => "&quot;",
      '<' => "&lt;",
      '>' => "&gt;",
    }

    # Escape first underscore of ST_Xstring sequences in input strings to appear as plaintext in Excel
    HEX_ESCAPE_REGEXP = /_(x[0-9A-Fa-f]{4}_)/
    XML_ESCAPE_UNDERSCORE = "_x005f_\\1"

    XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"

    WS_AROUND_TAGS = /(?<=>)\s+|\s+(?=<)/

    UNSAFE_ATTR_CHARS = /[&"<>]/
    UNSAFE_VALUE_CHARS = /[&<>]/

    # http://www.w3.org/TR/REC-xml/#NT-Char:
    # Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
    INVALID_XML10_CHARS = /[^\x09\x0A\x0D\x20-\x{D7FF}\x{E000}-\x{FFFD}\x{10000}-\x{10FFFF}]/

    # ST_Xstring escaping
    ESCAPE_CHAR = ->(c : Char) { "_x%04X_" % c.ord }

    def self.header : String
      XML_DECLARATION
    end

    def self.strip(xml : String) : String
      xml.gsub(WS_AROUND_TAGS, "")
    end

    def self.escape_attr(string : String) : String
      string.gsub(UNSAFE_ATTR_CHARS) { |c| XML_ESCAPES[c[0]] }
    end

    def self.escape_value(string : String) : String
      string
        .gsub(UNSAFE_VALUE_CHARS) { |c| XML_ESCAPES[c[0]] }
        .gsub(HEX_ESCAPE_REGEXP, XML_ESCAPE_UNDERSCORE)
        .gsub(INVALID_XML10_CHARS) { |m| ESCAPE_CHAR.call(m[0]) }
    end
  end
end
