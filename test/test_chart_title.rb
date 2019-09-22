# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestChartTitle < XlsxWriterTestCase
  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ].freeze

  test 'chart_title01' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
        chart.axis_id_1 = 46_165_376
        chart.axis_id_2 = 54_462_720

        series = chart.add_series('=Sheet1!$A$1:$A$5')
        series.name = 'Foo'
        chart.title = nil

        ws.insert_chart 'E9', chart
      end
    end
  end

  test 'chart_title02' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
        chart.axis_id_1 = 73_655_040
        chart.axis_id_2 = 73_656_576

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.title = 'Title!'

        ws.insert_chart 'E9', chart
      end
    end
  end
end
