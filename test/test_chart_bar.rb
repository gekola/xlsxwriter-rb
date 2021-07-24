# frozen_string_literal: true

require_relative './xlsx_func_testcase'
require_relative './support/chart_test'

class TestChartBar < XlsxWriterTestCase
  extend ChartTest

  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ].freeze

  chart_test 'chart_bar01', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 64_052_224
    chart.axis_id_2 = 64_055_552

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'
  end

  test 'chart_bar02' do |wb|
    wb.add_worksheet do |ws|
      ws.write_string('A1', 'Foo')
    end

    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

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
        chart.axis_id_1 = 64_265_216
        chart.axis_id_2 = 64_447_616

        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 86_048_128
        chart.axis_id_2 = 86_058_112

        chart.add_series '=Sheet1!$A$1:$A$4', '=Sheet1!$B$1:$B$4'
        chart.add_series '=Sheet1!$A$1:$A$4', '=Sheet1!$C$1:$C$4'

        ws.insert_chart 'F25', chart
      end
    end
  end

  test 'chart_bar04' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 64_446_848
        chart.axis_id_2 = 64_448_384

        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end
    end

    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 85_389_696
        chart.axis_id_2 = 85_391_232

        chart.add_series '=Sheet2!$A$1:$A$5', '=Sheet2!$B$1:$B$5'
        chart.add_series '=Sheet2!$A$1:$A$5', '=Sheet2!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end
    end
  end

  chart_test 'chart_bar05', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 64_264_064
    chart.axis_id_2 = 64_447_232

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'
  end

  chart_test 'chart_bar06', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 64_053_248
    chart.axis_id_2 = 64_446_464

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.name = 'Apple'
    chart.y_axis.name = 'Pear'
    chart.title = 'Title'
  end

  chart_test 'chart_bar08', XlsxWriter::Workbook::Chart::BAR do |chart, ws|
    ws.workbook.unset_default_url_format
    chart.axis_id_1 = 40_522_880
    chart.axis_id_2 = 40_524_416

    ws.write_url('A7', 'http://www.perl.com/')

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'
  end

  chart_test 'chart_bar09', XlsxWriter::Workbook::Chart::BAR_STACKED do |chart|
    chart.axis_id_1 = 40_274_560
    chart.axis_id_2 = 40_295_040

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'
  end

  chart_test 'chart_bar10', XlsxWriter::Workbook::Chart::BAR_STACKED_PERCENT do |chart|
    chart.axis_id_1 = 40_274_560
    chart.axis_id_2 = 40_295_040

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'
  end

  test 'chart_bar14' do |wb|
    wb.unset_default_url_format
    wb.add_worksheet

    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }
      ws.add_row ['http://www.perl.com/'], types: :url

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 40_294_272
        chart.axis_id_2 = 40_295_808

        chart.add_series '=Sheet2!$A$1:$A$5'
        chart.add_series '=Sheet2!$B$1:$B$5'
        chart.add_series '=Sheet2!$C$1:$C$5'

        ws.insert_chart 'E9', chart
      end

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 40_261_504
        chart.axis_id_2 = 65_749_760

        chart.add_series '=Sheet2!$A$1:$A$5'

        ws.insert_chart 'F25', chart
      end
    end

    wb.add_worksheet

    wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
      chart.axis_id_1 = 65_465_728
      chart.axis_id_2 = 66_388_352

      chart.add_series '=Sheet2!$A$1:$A$5'
      chart.add_series '=Sheet2!$B$1:$B$5'
      chart.add_series '=Sheet2!$C$1:$C$5'

      wb.add_chartsheet.chart = chart
    end
  end

  test 'chart_bar15' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 62_576_896
        chart.axis_id_2 = 62_582_784

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'

        wb.add_chartsheet.chart = chart
      end
    end

    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      wb.add_chart(XlsxWriter::Workbook::Chart::COLUMN) do |chart|
        chart.axis_id_1 = 65_979_904
        chart.axis_id_2 = 65_981_440

        chart.add_series '=Sheet2!$A$1:$A$5'

        wb.add_chartsheet.chart = chart
      end
    end
  end

  chart_test 'chart_bar16', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 64_784_640
    chart.axis_id_2 = 65_429_504

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.title = 'Title'
    chart.x_axis.name = 'Apple'
    chart.y_axis.name = 'Pear'
    chart.legend_position = XlsxWriter::Workbook::Chart::LEGEND_BOTTOM

    chart.workbook.add_chartsheet.chart = chart
  end

  chart_test 'chart_bar17', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 40_294_272
    chart.axis_id_2 = 40_295_808

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    cs = chart.workbook.add_chartsheet
    cs.activate
    cs.chart = chart
  end

  chart_test 'chart_bar18', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 61_298_176
    chart.axis_id_2 = 61_300_096

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.workbook.add_chartsheet do |cs|
      cs.set_margins(0.7086614173228347, 0.7086614173228347, 0.7480314960629921, 0.7480314960629921)
      cs.paper = 9
      cs.set_header('Page &P', margin: 0.3149606299212598)
      cs.set_footer('&A', margin: 0.3149606299212598)
      cs.activate
      cs.horizontal_dpi = 200
      cs.vertical_dpi = 200
      cs.chart = chart
    end
  end

  chart_test 'chart_bar19', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 66_558_592
    chart.axis_id_2 = 66_569_344

    chart.add_series.set_values 'Sheet1', 0, 0, 4, 0
    chart.add_series.set_values 'Sheet1', 0, 1, 4, 1
    chart.add_series.set_values 'Sheet1', 0, 2, 4, 2

    chart.x_axis.name = '=Sheet1!$A$2'
    chart.y_axis.name = '=Sheet1!$A$3'
    chart.title = '=Sheet1!$A$1'
  end

  chart_test 'chart_bar20', XlsxWriter::Workbook::Chart::BAR do |chart, ws|
    chart.axis_id_1 = 45_925_120
    chart.axis_id_2 = 45_927_040

    ws.add_row []
    ws.add_row ['Pear']

    chart.add_series('=Sheet1!$A$1:$A$5')
    chart.add_series('=Sheet1!$B$1:$B$5').name = 'Apple'
    chart.add_series('=Sheet1!$C$1:$C$5').name = '=Sheet1!$A$7'
  end

  chart_test 'chart_bar21', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 64_052_224
    chart.axis_id_2 = 64_055_552

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'
  end

  test 'chart_bar22' do |wb|
    wb.add_worksheet do |ws|
      [
        [nil, *(1..3).map { |i| "Series #{i}" }],
        ['Category 1', 4.3, 2.4, 2],
        ['Category 2', 2.5, 4.5, 2],
        ['Category 3', 3.5, 1.8, 3],
        ['Category 4', 4.5, 2.8, 5]
      ].each { |row| ws.add_row row }
      ws.set_column('A', 'D', width: 11)

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 43_706_240
        chart.axis_id_2 = 43_727_104

        chart.add_series '=Sheet1!$A$2:$A$5', '=Sheet1!$B$2:$B$5'
        chart.add_series '=Sheet1!$A$2:$A$5', '=Sheet1!$C$2:$C$5'
        chart.add_series '=Sheet1!$A$2:$A$5', '=Sheet1!$D$2:$D$5'

        ws.insert_chart 'E9', chart
      end
    end
  end

  # tests 51 - 55 use +lxw_chart_add_data_cache+ which is not exposed
  # in this library

  chart_test 'chart_bar61', XlsxWriter::Workbook::Chart::BAR, ref_file_name: 'chart_bar01' do |chart|
    chart.axis_id_1 = 64_052_224
    chart.axis_id_2 = 64_055_552

    series1 = chart.add_series
    series2 = chart.add_series

    series1.set_categories 'Sheet1', 0, 0, 4, 0
    series1.set_values     'Sheet1', 0, 1, 4, 1

    series2.set_categories 'Sheet1', 0, 0, 4, 0
    series2.set_values     'Sheet1', 0, 2, 4, 2
  end

  chart_test 'chart_bar65', XlsxWriter::Workbook::Chart::BAR, ref_file_name: 'chart_bar05' do |chart|
    chart.axis_id_1 = 64_264_064
    chart.axis_id_2 = 64_447_232

    chart.add_series.set_values 'Sheet1', 0, 0, 4, 0
    chart.add_series.set_values 'Sheet1', 0, 1, 4, 1
    chart.add_series.set_values 'Sheet1', 0, 2, 4, 2
  end

  chart_test 'chart_bar69', XlsxWriter::Workbook::Chart::BAR, ref_file_name: 'chart_bar19' do |chart|
    chart.axis_id_1 = 66_558_592
    chart.axis_id_2 = 66_569_344

    chart.add_series.set_values 'Sheet1', 'A1:A5'
    chart.add_series.set_values 'Sheet1', 'B1', 'B5'
    chart.add_series.set_values 'Sheet1', 0, 2, 4, 2

    chart.x_axis.set_name_range 'Sheet1', 'A2'
    chart.y_axis.set_name_range 'Sheet1', 'A3'
    chart.set_name_range 'Sheet1', 'A1'
  end

  chart_test 'chart_bar70', XlsxWriter::Workbook::Chart::BAR, ref_file_name: 'chart_bar20' do |chart, ws|
    chart.axis_id_1 = 45_925_120
    chart.axis_id_2 = 45_927_040

    ws.add_row []
    ws.add_row ['Pear']

    chart.add_series.set_values('Sheet1', 'A1:A5')
    chart.add_series.set_values('Sheet1', 'B1:B5').name = 'Apple'
    chart.add_series.set_values('Sheet1', 'C1:C5').name = '=Sheet1!$A$7'
  end
end
