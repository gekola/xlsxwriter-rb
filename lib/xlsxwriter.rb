module XlsxWriter
  VERSION='0.0.1'.freeze
  class Workbook
    def self.open(file, options={})
      wb = new(file, options)
      if block_given?
        yield wb
        wb.free
        nil
      else
        wb
      end
    end
  end
end

require 'xlsxwriter/xlsxwriter'
require 'xlsxwriter/worksheet'
