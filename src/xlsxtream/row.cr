require "./xml"

module Xlsxtream
  alias CellValue = Nil | Bool | Int8 | Int16 | Int32 | Int64 | Int128 | UInt8 | UInt16 | UInt32 | UInt64 | UInt128 | Float32 | Float64 | String | Time | Date

  class Row
    NUMBER_PATTERN  = /\A-?[0-9]+(\.[0-9]+)?\z/
    # ISO 8601 yyyy-mm-dd
    DATE_PATTERN    = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/
    # ISO 8601 yyyy-mm-ddThh:mm:ss(.s)(Z|+hh:mm|-hh:mm)
    TIME_PATTERN    = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/

    TRUE_STRING  = "true"
    FALSE_STRING = "false"

    def initialize(@row : Array(CellValue), @rownum : Int32,
                   @sst : SharedStringTable? = nil,
                   @auto_format : Bool = false,
                   @is_header : Bool = false)
    end

    def to_xml : String
      col_index = 0
      xml = String::Builder.new
      xml << %(<row r="#{@rownum}">)

      if @is_header
        normal_style = %( s="3")
        date_style   = %( s="4")
        time_style   = %( s="5")
      else
        normal_style = ""
        date_style   = %( s="1")
        time_style   = %( s="2")
      end

      @row.each do |value|
        cid = "#{column_name(col_index)}#{@rownum}"
        col_index += 1

        if @auto_format && value.is_a?(String)
          value = auto_format(value)
        end

        case value
        when Number
          xml << %(<c r="#{cid}"#{normal_style} t="n"><v>#{value}</v></c>)
        when Bool
          xml << %(<c r="#{cid}"#{normal_style} t="b"><v>#{value ? 1 : 0}</v></c>)
        when Time
          xml << %(<c r="#{cid}"#{time_style}><v>#{time_to_oa_date(value)}</v></c>)
        when Date
          xml << %(<c r="#{cid}"#{date_style}><v>#{date_to_oa_date(value)}</v></c>)
        when String
          unless value.empty?
            if sst = @sst
              xml << %(<c r="#{cid}"#{normal_style} t="s"><v>#{sst[value]}</v></c>)
            else
              xml << %(<c r="#{cid}"#{normal_style} t="inlineStr"><is><t>#{XML.escape_value(value)}</t></is></c>)
            end
          end
        when Nil
          # no xml output for nil
        end
      end

      xml << "</row>"
      xml.to_s
    end

    private def auto_format(value : String) : CellValue
      case value
      when TRUE_STRING
        true
      when FALSE_STRING
        false
      when NUMBER_PATTERN
        value.includes?('.') ? value.to_f : value.to_i
      when DATE_PATTERN
        parse_date(value)
      when TIME_PATTERN
        parse_time(value)
      else
        value
      end
    end

    private def parse_date(value : String) : CellValue
      parts = value.split('-')
      year = parts[0].to_i
      month = parts[1].to_i
      day = parts[2].to_i
      # Validate the date
      Time.utc(year, month, day)
      Date.new(year, month, day)
    rescue
      value
    end

    private def parse_time(value : String) : CellValue
      # Try various ISO 8601 formats
      # Crystal's Time.parse_iso8601 handles most cases but needs at least seconds
      # We need to handle the case without seconds (e.g., "1900-01-01T12:00")
      normalized = value
      # If no seconds, add :00
      if normalized =~ /T\d{2}:\d{2}(?:Z|[+-]\d{2}:\d{2})?$/
        # Insert :00 before timezone
        if normalized.ends_with?('Z')
          normalized = normalized[0..-2] + ":00Z"
        elsif normalized =~ /([+-]\d{2}:\d{2})$/
          tz = $1
          normalized = normalized[0..-(tz.size + 1)] + ":00" + tz
        else
          normalized = normalized + ":00"
        end
      end
      # If no timezone, treat as UTC
      unless normalized =~ /Z$|[+-]\d{2}:\d{2}$/
        normalized = normalized + "Z"
      end
      time = Time.parse_iso8601(normalized)
      # Convert to UTC-equivalent for OA date calculation
      time
    rescue
      value
    end

    # Converts Time to OLE Automation Date
    # Local dates are stored as UTC by truncating the offset
    private def time_to_oa_date(time : Time) : Float64
      (time.to_unix_f + time.offset.to_f) / 86400.0 + 25569.0
    end

    # Converts Date to OLE Automation Date
    private def date_to_oa_date(date : Date) : Float64
      jd = julian_day(date.year, date.month, date.day)
      (jd - 2415019).to_f
    end

    private def julian_day(year : Int32, month : Int32, day : Int32) : Int64
      a = ((14 - month) / 12).to_i64
      y = year.to_i64 + 4800 - a
      m = month.to_i64 + 12 * a - 3
      day.to_i64 + ((153 * m + 2) / 5).to_i64 + 365 * y + (y / 4).to_i64 - (y / 100).to_i64 + (y / 400).to_i64 - 32045
    end

    private def column_name(index : Int32) : String
      name = ""
      n = index + 1
      while n > 0
        n -= 1
        name = (('A'.ord + n % 26).chr) + name
        n //= 26
      end
      name
    end
  end
end
