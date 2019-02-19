# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestChartScatter < XlsxWriterTestCase
  test 'chart_scatter15' do |wb|
    wb.add_worksheet do |ws|
      wb.add_chart(XlsxWriter::Workbook::Chart::SCATTER) do |chart|
        chart.axis_id_1 = 58_843_520
        chart.axis_id_2 = 58_845_440


        ws.write_string 0, 0, 'X'
        ws.write_string 0, 1, 'Y'
        ws.write_number 1, 0,  1
        ws.write_number 1, 1, 10
        ws.write_number 2, 0,  3
        ws.write_number 2, 1, 30

        chart.add_series '=Sheet1!$A$2:$A$3', '=Sheet1!$B$2:$B$3'

        chart.x_axis.set_name_range 'Sheet1', 0, 0
        chart.x_axis.set_name_font italic: true, baseline: -1

        chart.y_axis.set_name_range 'Sheet1', 0, 1

        ws.insert_chart 'E9', chart
      end
    end
  end
end
