require 'xlsxwriter'
require_relative './support/xlsx_comparable'

class XlsxWriterTestCase < Test::Unit::TestCase
  include XlsxComparable

  def self.test(name, &block)
    define_method :"test_#{name}" do
      file_path = "tmp/#{name}.xlsx"
      ref_file_path = "ext/xlsxwriter/libxlsxwriter/test/functional/xlsx_files/#{name}.xlsx"
      XlsxWriter::Workbook.open(file_path) do |wb|
        yield wb
      end

      assert_xlsx_equal file_path, ref_file_path
    end
  end
end
