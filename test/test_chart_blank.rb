# frozen_string_literal: true

require_relative './xlsx_func_testcase'
require_relative './support/chart_test'

class TestChartBlank < XlsxWriterTestCase
  extend ChartTest

  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ].freeze

  chart_test 'chart_blank01', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 45_705_856
    chart.axis_id_2 = 45_843_584

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.show_blank_as = XlsxWriter::Workbook::Chart::BLANKS_AS_GAP
  end
end
