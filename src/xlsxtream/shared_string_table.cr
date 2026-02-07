module Xlsxtream
  class SharedStringTable
    @data = Hash(String, Int32).new
    @references = 0

    def [](string : String) : Int32
      @references += 1
      @data.put_if_absent(string) { @data.size }
    end

    def references : Int32
      @references
    end

    def size : Int32
      @data.size
    end

    def empty? : Bool
      @data.empty?
    end

    def each_key(& : String ->) : Nil
      @data.each_key { |k| yield k }
    end
  end
end
