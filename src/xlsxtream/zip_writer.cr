require "compress/zip"

module Xlsxtream
  class ZipWriter
    BUFFER_SIZE = 64 * 1024

    alias Closeable = Compress::Zip::Writer | IO

    def self.with_output_to(output : ZipWriter) : ZipWriter
      output
    end

    def self.with_output_to(output : String | Path) : ZipWriter
      file = File.open(output.to_s, "wb")
      zip = Compress::Zip::Writer.new(file)
      new(zip, close: [zip.as(Closeable), file.as(Closeable)])
    end

    def self.with_output_to(output : IO) : ZipWriter
      zip = Compress::Zip::Writer.new(output)
      new(zip, close: [zip.as(Closeable)])
    end

    @current_writer : IO? = nil
    @entry_done : Channel(Nil)? = nil
    @buffer : String::Builder = String::Builder.new

    def initialize(@zip : Compress::Zip::Writer, @close : Array(Closeable) = [] of Closeable)
    end

    def <<(data : String) : self
      @buffer << data
      flush_buffer if @buffer.bytesize >= BUFFER_SIZE
      self
    end

    def add_file(path : String) : Nil
      flush_file
      reader, @current_writer = IO.pipe
      @entry_done = Channel(Nil).new
      spawn do
        @zip.add(path, reader)
        reader.close
        @entry_done.not_nil!.send(nil)
      end
    end

    def close : Nil
      flush_file
      @close.each(&.close)
    end

    private def flush_buffer : Nil
      @current_writer.not_nil!.print(@buffer.to_s)
      @buffer = String::Builder.new
    end

    private def flush_file : Nil
      return unless (writer = @current_writer)
      flush_buffer if @buffer.bytesize > 0
      writer.close
      @entry_done.not_nil!.receive
      @current_writer = nil
    end
  end
end
