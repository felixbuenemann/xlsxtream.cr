require "../spec_helper"
require "compress/zip"

module Xlsxtream
  describe ZipWriter do
    it "writes of multiple files" do
      tempfile = File.tempfile("ztio-test")
      begin
        io = Xlsxtream::ZipWriter.with_output_to(tempfile.path)
        io.add_file("book1.xml")
        io << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook />"
        io.add_file("book2.xml")
        io << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook>"
        io << "</workbook>"
        io.add_file("empty.txt")
        io.add_file("another.xml")
        io << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><another />"
        io.close

        file_contents = Hash(String, String).new
        Compress::Zip::File.open(tempfile.path) do |zip_file|
          zip_file.entries.each do |entry|
            entry.open do |entry_io|
              file_contents[entry.filename] = entry_io.gets_to_end
            end
          end
        end
        file_contents["empty.txt"].should eq("")
        file_contents["book2.xml"].should eq("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook></workbook>")
        file_contents["another.xml"].should eq("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><another />")
      ensure
        tempfile.delete
      end
    end

    it "with_output_to wraps another writer" do
      tempfile = File.tempfile("ztio-test")
      begin
        zip = Compress::Zip::Writer.new(File.open(tempfile.path, "wb"))
        another_writer = Xlsxtream::ZipWriter.new(zip)
        Xlsxtream::ZipWriter.with_output_to(another_writer).should eq(another_writer)
        another_writer.close
      ensure
        tempfile.delete
      end
    end

    it "with_output_to creates a file with a given path" do
      tempfile = File.tempfile("xlsxtream-output", ".xlsx")
      begin
        path = tempfile.path
        tempfile.delete # remove so with_output_to can create it
        writer = Xlsxtream::ZipWriter.with_output_to(path)
        File.exists?(path).should be_true
        writer.close
        File.exists?(path).should be_true
      ensure
        File.delete?(path.not_nil!) if path
      end
    end

    it "with_output_to writes into io" do
      io = IO::Memory.new
      writer = Xlsxtream::ZipWriter.with_output_to(io)
      writer.close
      io.size.should be > 0
    end
  end
end
