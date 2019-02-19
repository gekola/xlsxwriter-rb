# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestChartBar < XlsxWriterTestCase
  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ]

  # 61, 01 does not add any coverage
  test 'chart_bar01' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 64_052_224
        chart.axis_id_2 = 64_055_552

        series1 = chart.add_series
        series2 = chart.add_series

        series1.set_categories 'Sheet1', 0, 0, 4, 0
        series1.set_values     'Sheet1', 0, 1, 4, 1

        series2.set_categories 'Sheet1', 0, 0, 4, 0
        series2.set_values     'Sheet1', 0, 2, 4, 2

        ws.insert_chart 'E9', chart
      end
    end
  end
end
