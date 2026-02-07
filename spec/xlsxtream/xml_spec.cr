require "../spec_helper"

module Xlsxtream
  describe XML do
    it "header" do
      expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      XML.header.should eq(expected)
    end

    it "strip" do
      xml = <<-XML
        <hello id="1">
          <world/>
        </hello>
      XML
      expected = "<hello id=\"1\"><world/></hello>"
      XML.strip(xml).should eq(expected)
    end

    it "escape_attr" do
      unsafe_attribute = "<hello> & \"world\""
      expected = "&lt;hello&gt; &amp; &quot;world&quot;"
      XML.escape_attr(unsafe_attribute).should eq(expected)
    end

    it "escape_value" do
      unsafe_value = "<hello> & \"world\""
      expected = "&lt;hello&gt; &amp; \"world\""
      XML.escape_value(unsafe_value).should eq(expected)
    end

    it "escape_value invalid xml chars" do
      unsafe_value = "The \u{07} rings\u{00}\u{FFFE}\u{FFFF}"
      expected = "The _x0007_ rings_x0000__xFFFE__xFFFF_"
      XML.escape_value(unsafe_value).should eq(expected)
    end

    it "escape_value valid xml chars" do
      safe_value = "\u{10000}\u{10FFFF}"
      expected = safe_value
      XML.escape_value(safe_value).should eq(expected)
    end

    it "encode underscores using x005f" do
      unsafe_value = "_xDcc2_"
      safe_value = "_x005f_xDcc2_"
      XML.escape_value(unsafe_value).should eq(safe_value)
    end

    it "encode underscores using x005f multiple occurrences" do
      unsafe_value = "_xDcc2_aa_x3d12_bb_xDea3_cc_xDaa5_"
      safe_value = "_x005f_xDcc2_aa_x005f_x3d12_bb_x005f_xDea3_cc_x005f_xDaa5_"
      XML.escape_value(unsafe_value).should eq(safe_value)
    end

    it "not escaping regular underscores" do
      safe_value = "this_test_does_not_replace_underscores_xDcc2"
      XML.escape_value(safe_value).should eq(safe_value)
    end
  end
end
