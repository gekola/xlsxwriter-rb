# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestChartLegend < XlsxWriterTestCase
  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ].freeze

  test 'chart_legend01' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
        chart.axis_id_1 = 54_461_952
        chart.axis_id_2 = 54_463_872

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        chart.legend_position = XlsxWriter::Workbook::Chart::LEGEND_NONE

        ws.insert_chart 'E9', chart
      end
    end
  end

  test 'chart_legend03' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::LINE) do |chart|
        chart.axis_id_1 = 93_548_928
        chart.axis_id_2 = 93_550_464

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        chart.legend_position = XlsxWriter::Workbook::Chart::LEGEND_TOP_RIGHT

        ws.insert_chart 'E9', chart
      end
    end
  end

  test 'chart_legend04' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::LINE) do |chart|
        chart.axis_id_1 = 93_548_928
        chart.axis_id_2 = 93_550_464

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        chart.legend_position = XlsxWriter::Workbook::Chart::LEGEND_OVERLAY_TOP_RIGHT

        ws.insert_chart 'E9', chart
      end
    end
  end
end
