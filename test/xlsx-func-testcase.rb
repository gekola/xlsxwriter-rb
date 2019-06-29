# frozen_string_literal: true

require 'xlsxwriter'
require_relative './support/xlsx_comparable'
require_relative './support/with_xlsx_file'

class XlsxWriterTestCaseConfig
  attr_accessor :ignore_elements, :ignore_files
  def initialize
    @ignore_elements = {}
    @ignore_files = []
  end
end

class XlsxWriterTestCase < Test::Unit::TestCase
  include WithXlsxFile
  include XlsxComparable

  def self.test(name, **opts, &block)
    define_method(:"test_#{name}") do
      file_path = "tmp/#{name}.xlsx"
      ref_name = opts && opts.delete(:ref_file_name) || name
      ref_file_path = "ext/xlsxwriter/libxlsxwriter/test/functional/xlsx_files/#{ref_name}"
      ref_file_path += '.xlsx' if File.extname(ref_file_path) == ''
      tc = XlsxWriterTestCaseConfig.new

      compare_files = proc {
        assert_xlsx_equal file_path, ref_file_path, tc.ignore_files, tc.ignore_elements
      }

      with_xlsx_file(file_path, **opts, after: compare_files) do |wb|
        instance_exec(wb, tc, &block)
      end
    end
  end

  def image_path(path)
    File.expand_path(File.join('../../ext/xlsxwriter/libxlsxwriter/test/functional/src/images', path), __FILE__)
  end
end
