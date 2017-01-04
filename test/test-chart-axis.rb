require_relative './xlsx-func-testcase'

class TestChartAxis < XlsxWriterTestCase
  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ]

  test 'chart_axis01' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 58_955_648
        chart.axis_id_2 = 68_446_848

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'

        chart.x_axis.name = 'XXX'
        chart.y_axis.name = 'YYY'

        ws.insert_chart 'E9', chart
      end
    end
  end

  test 'chart_axis06' do |wb|
    wb.add_worksheet do |ws|
      [ [2, 60], [4, 30], [6, 10] ].each { |row| ws.add_row row }
      wb.add_chart(XlsxWriter::Workbook::Chart::PIE) do |chart|
        chart.add_series '=Sheet1!$A$1:$A$3', '=Sheet1!$B$1:$B$3'
        chart.title = 'Title'
        chart.x_axis.name = 'XXX'
        chart.y_axis.name = 'YYY'
        ws.insert_chart('E9', chart)
      end
    end
  end

  test 'chart_axis26' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::LINE) do |chart|
        chart.axis_id_1 = 73_048_448
        chart.axis_id_2 = 73_049_984

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'

        chart.x_axis.set_num_font rotation: 45, baseline: -1

        ws.insert_chart('E9', chart)
      end
    end
  end

  test 'chart_axis35' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
        chart.axis_id_1 = 63_008_128
        chart.axis_id_2 = 62_522_496

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'

        chart.y_axis.set_line none: true

        ws.insert_chart('E9', chart)
      end
    end
  end

  test 'chart_axis38' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
        chart.axis_id_1 = 45_642_496
        chart.axis_id_2 = 45_644_416

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'

        chart.y_axis.set_line color: XlsxWriter::Format::COLOR_YELLOW
        chart.y_axis.set_fill color: XlsxWriter::Format::COLOR_RED

        ws.insert_chart 'E9', chart
      end
    end
  end
end
