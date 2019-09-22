# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestChartPie < XlsxWriterTestCase
  DATA = [[2, 60], [4, 30], [6, 10]].freeze

  test 'chart_pie02' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::PIE) do |chart|
        chart.add_series '=Sheet1!$A$1:$A$3', '=Sheet1!$B$1:$B$3'
        chart.legend_set_font bold: true, italic: true, baseline: -1

        ws.insert_chart('E9', chart)
      end
    end
  end

  test 'chart_pie03' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::PIE) do |chart|
        chart.add_series '=Sheet1!$A$1:$A$3', '=Sheet1!$B$1:$B$3'
        chart.legend_delete_series [1]

        ws.insert_chart('E9', chart)
      end
    end
  end

  test 'chart_pie04' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::PIE) do |chart|
        chart.add_series '=Sheet1!$A$1:$A$3', '=Sheet1!$B$1:$B$3'
        chart.legend_position = XlsxWriter::Workbook::Chart::LEGEND_OVERLAY_RIGHT

        ws.insert_chart('E9', chart)
      end
    end
  end
end
