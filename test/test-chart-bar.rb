# frozen_string_literal: true

require_relative './xlsx-func-testcase'
require_relative './support/chart_test'

class TestChartBar < XlsxWriterTestCase
  extend ChartTest

  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ]

  chart_test 'chart_bar01', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 64_052_224
    chart.axis_id_2 = 64_055_552

    series1 = chart.add_series
    series2 = chart.add_series

    series1.set_categories 'Sheet1', 0, 0, 4, 0
    series1.set_values     'Sheet1', 0, 1, 4, 1

    series2.set_categories 'Sheet1', 0, 0, 4, 0
    series2.set_values     'Sheet1', 0, 2, 4, 2
  end

  test 'chart_bar02' do |wb|
    ws1 = wb.add_worksheet
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }
      ws1.write_string('A1', 'Foo')

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 93_218_304
        chart.axis_id_2 = 93_219_840

        chart.add_series '=Sheet2!$A$1:$A$5', '=Sheet2!$B$1:$B$5'
        chart.add_series '=Sheet2!$A$1:$A$5', '=Sheet2!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end
    end
  end

  test 'chart_bar03' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 64265216
        chart.axis_id_2 = 64447616

        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 86048128
        chart.axis_id_2 = 86058112

        chart.add_series '=Sheet1!$A$1:$A$4', '=Sheet1!$B$1:$B$4'
        chart.add_series '=Sheet1!$A$1:$A$4', '=Sheet1!$C$1:$C$4'

        ws.insert_chart 'F25', chart
      end
    end
  end
end
