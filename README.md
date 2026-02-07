# Xlsxtream

[![CI](https://github.com/felixbuenemann/xlsxtream.cr/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/felixbuenemann/xlsxtream.cr/actions/workflows/ci.yml)

This is a Crystal port of the [xlsxtream](https://github.com/felixbuenemann/xlsxtream) Ruby gem.

Xlsxtream is a streaming writer for XLSX spreadsheets. It supports multiple worksheets and optional string
deduplication via a shared string table (SST). Its purpose is to replace CSV for large exports, because using
CSV in Excel is very buggy and error prone. It's very efficient and can quickly write millions of rows with
low memory usage.

Xlsxtream does not support formatting, charts, comments and a myriad of
other [OOXML](https://en.wikipedia.org/wiki/Office_Open_XML) features.

Xlsxtream supports writing to files or IO objects, data is flushed as the ZIP compressor sees fit.

## Installation

Add this to your application's `shard.yml`:

```yaml
dependencies:
  xlsxtream:
    github: felixbuenemann/xlsxtream.cr
```

## Usage

```crystal
require "xlsxtream"

# Creates a new workbook file, write and close it at the end of the block
Xlsxtream::Workbook.open("my_data.xlsx") do |xlsx|
  xlsx.write_worksheet("Sheet1") do |sheet|
    # Bool, Xlsxtream::Date, Time, Int32, Int64, Float64 and String are properly mapped
    sheet << [true, Xlsxtream::Date.new(2024, 1, 1), "hello", "world", 42, 3.14159265359, 42_i64 ** 13] of Xlsxtream::CellValue
  end
end

io = IO::Memory.new
xlsx = Xlsxtream::Workbook.new(io)

# Number of columns doesn't have to match
xlsx.write_worksheet("Sheet1") do |sheet|
  sheet << ["first", "row"] of Xlsxtream::CellValue
  sheet << ["second", "row", "with", "more columns"] of Xlsxtream::CellValue
end

# Write multiple worksheets with custom names
xlsx.write_worksheet("AppendixSheet") do |sheet|
  sheet.add_row ["Timestamp", "Comment"] of Xlsxtream::CellValue
  sheet.add_row [Time.utc, "Good times"] of Xlsxtream::CellValue
  sheet.add_row [Time.utc, "Time-machine"] of Xlsxtream::CellValue
end

# If you have highly repetitive data, you can enable Shared String Tables (SST)
# for the workbook or a single worksheet. The SST has to be kept in memory,
# so do not use it if you have a huge amount of rows or a little duplication
# of content across cells. A single SST is used for the whole workbook.
xlsx.write_worksheet(name: "SheetWithSST", use_shared_strings: true) do |sheet|
  sheet << ["the", "same", "old", "story"] of Xlsxtream::CellValue
  sheet << ["the", "old", "same", "story"] of Xlsxtream::CellValue
  sheet << ["old", "the", "same", "story"] of Xlsxtream::CellValue
end

# Strings in numeric or date/time format can be auto-detected and formatted
# appropriately. This is a convenient way to avoid an Excel-warning about
# "Number stored as text". Dates and times must be in the ISO-8601 format and
# numeric values must contain only numbers and an optional decimal separator.
# The strings true and false are detected as boolean values.
xlsx.write_worksheet(name: "SheetWithAutoFormat", auto_format: true) do |sheet|
  # these two rows will be identical in the xlsx-output
  sheet << [true, 11.85, Time.utc(2050, 1, 1, 12, 0, 0), Xlsxtream::Date.new(1984, 1, 1)] of Xlsxtream::CellValue
  sheet << ["true", "11.85", "2050-01-01T12:00:00Z", "1984-01-01"] of Xlsxtream::CellValue
end

# You can also create worksheet without a block, using the `add_worksheet` method.
# It can only be used sequentially, so remember to manually close the worksheet
# when you are done (before opening a new one).
worksheet = xlsx.add_worksheet(name: "SheetWithoutBlock")
worksheet << ["some", "data"] of Xlsxtream::CellValue
worksheet.close

# Writes metadata and ZIP archive central directory
xlsx.close

# Changing the default font from Calibri, 12pt, Swiss
Xlsxtream::Workbook.new(io,
  font: Xlsxtream::FontOptions.new(
    name: "Times New Roman",
    size: 10, # size in pt
    family: "Roman" # Swiss, Modern, Script, Decorative
  )
)

# Treat the first output row as a header, using bold and centred text
Xlsxtream::Workbook.new(io, has_header_row: true)

# Specifying column widths in pixels or characters; 3 column example;
# "pixel" widths appear to be *relative* to an assumed 11pt Calibri
# font, so if selecting a different font or size (see above), do not
# adjust widths to match. Calculate pixel widths for 11pt Calibri.
Xlsxtream::Workbook.new(io, columns: [
  {width_chars: nil, width_pixels: 33.0},
  {width_chars: 7, width_pixels: nil},
  {width_chars: 24, width_pixels: nil},
])
# The columns option can also be given to write_worksheet, so it's
# possible to have multiple worksheets with different column widths.
```

## Development

After checking out the repo, run `crystal spec` to run the tests.

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/felixbuenemann/xlsxtream.cr.

## License

The shard is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).
