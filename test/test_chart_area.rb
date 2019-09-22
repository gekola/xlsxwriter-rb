# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestChartArea < XlsxWriterTestCase
  DATA = [
    [1,  8,  3],
    [2,  7,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ].freeze

  test 'chart_area01' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::AREA) do |chart|
        chart.axis_id_1 = 43_407_616
        chart.axis_id_2 = 43_433_984

        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end
    end
  end

  test 'chart_area02' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::AREA_STACKED) do |chart|
        chart.axis_id_1 = 62_813_312
        chart.axis_id_2 = 62_814_848

        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end
    end
  end

  test 'chart_area03' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::AREA_STACKED_PERCENT) do |chart|
        chart.axis_id_1 = 62_813_312
        chart.axis_id_2 = 62_814_848

        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end
    end
  end
end
