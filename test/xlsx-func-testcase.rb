# frozen_string_literal: true

require 'xlsxwriter'
require_relative './support/xlsx_comparable'

class XlsxWriterTestCaseConfig
  attr_accessor :ignore_elements, :ignore_files
  def initialize
    @ignore_elements = {}
    @ignore_files = []
  end
end

class XlsxWriterTestCase < Test::Unit::TestCase
  include XlsxComparable

  def self.test(name, opts=nil, &block)
    define_method(:"test_#{name}") do
      file_path = "tmp/#{name}.xlsx"
      ref_name = opts && opts.delete(:ref_file_name) || name
      ref_file_path = "ext/xlsxwriter/libxlsxwriter/test/functional/xlsx_files/#{ref_name}.xlsx"
      begin
        tc = XlsxWriterTestCaseConfig.new
        args = [file_path, opts].compact
        XlsxWriter::Workbook.open(*args) do |wb|
          instance_exec(wb, tc, &block)
        end

        assert_xlsx_equal file_path, ref_file_path, tc.ignore_files, tc.ignore_elements
      ensure
        File.delete file_path if File.exist? file_path
      end
    end
  end

  def image_path(path)
    File.expand_path(File.join('../../ext/xlsxwriter/libxlsxwriter/test/functional/src/images', path), __FILE__)
  end
end
