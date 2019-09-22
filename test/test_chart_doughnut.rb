# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestChartDoughnut < XlsxWriterTestCase
  DATA = [[2, 60], [4, 30], [6, 10]].freeze

  test 'chart_doughnut02' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::DOUGHNUT) do |chart|
        chart.add_series '=Sheet1!$A$1:$A$3', '=Sheet1!$B$1:$B$3'
        chart.hole_size = 10

        ws.insert_chart('E9', chart)
      end
    end
  end

  test 'chart_doughnut04' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::DOUGHNUT) do |chart|
        chart.add_series '=Sheet1!$A$1:$A$3', '=Sheet1!$B$1:$B$3'
        chart.rotation = 30

        ws.insert_chart('E9', chart)
      end
    end
  end
end
