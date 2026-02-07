module Xlsxtream
  struct Date
    getter year : Int32
    getter month : Int32
    getter day : Int32

    def initialize(@year : Int32, @month : Int32, @day : Int32)
    end
  end
end
