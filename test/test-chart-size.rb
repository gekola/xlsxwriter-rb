# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestChartSize < XlsxWriterTestCase
  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ]

  test 'chart_size01' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
        chart.axis_id_1 = 61_355_904
        chart.axis_id_2 = 61_365_248

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'

        ws.insert_chart('E9', chart, x_scale: 1.06666667, y_scale: 1.11111112)
      end
    end
  end

  test 'chart_size04' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
        chart.axis_id_1 = 73_773_440
        chart.axis_id_2 = 73_774_976

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'

        ws.insert_chart('E9', chart, x_offset: 8, y_offset: 9)
      end
    end
  end
end
