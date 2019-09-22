# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestChartsheet < XlsxWriterTestCase
  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15]
  ].freeze

  test 'chartsheet01' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.activate

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 79_858_304
        chart.axis_id_2 = 79_860_096

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        cs.chart = chart
      end
    end
  end

  test 'chartsheet02' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.activate

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 79_858_304
        chart.axis_id_2 = 79_860_096

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        cs.chart = chart
      end
    end
    wb.add_worksheet.select
  end

  test 'chartsheet03' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.hide

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 43_794_816
        chart.axis_id_2 = 43_796_352

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        cs.chart = chart
      end
    end
  end

  test 'chartsheet04' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.protect

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 43_913_216
        chart.axis_id_2 = 43_914_752

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        cs.chart = chart
      end
    end
  end

  test 'chartsheet05' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.zoom = 75

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 43_695_104
        chart.axis_id_2 = 43_787_008

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        cs.chart = chart
      end
    end
  end

  test 'chartsheet06' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.tab_color = XlsxWriter::Format::COLOR_RED

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 43_778_432
        chart.axis_id_2 = 43_780_352

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        cs.chart = chart
      end
    end
  end

  test 'chartsheet07' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.set_portrait
      cs.paper = 9
      cs.horizontal_dpi = cs.vertical_dpi = 200

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 61_296_640
        chart.axis_id_2 = 61_298_176

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        cs.chart = chart
      end
    end
  end

  test 'chartsheet08' do |wb, t|
    t.ignore_files = ['xl/drawings/drawing1.xml']
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.set_margins(
        0.511811023622047,
        0.511811023622047,
        0.551181102362204,
        0.944881889763779
      )
      cs.paper = 9
      cs.set_portrait
      cs.set_header('&CPage &P', margin: 0.118110236220472)
      cs.set_footer('&C&A', margin: 0.118110236220472)
      cs.horizontal_dpi = cs.vertical_dpi = 200

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 61_297_792
        chart.axis_id_2 = 61_299_328

        chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'
        cs.chart = chart
      end
    end
  end

  test 'chartsheet09' do |wb|
    wb.add_worksheet do |ws|
      DATA.each { |row| ws.add_row row }

      cs = wb.add_chartsheet
      cs.activate

      wb.add_chart(XlsxWriter::Workbook::Chart::BAR) do |chart|
        chart.axis_id_1 = 49_044_480
        chart.axis_id_2 = 49_055_232

        series = chart.add_series '=Sheet1!$A$1:$A$5'
        chart.add_series '=Sheet1!$B$1:$B$5'
        chart.add_series '=Sheet1!$C$1:$C$5'

        series.set_line(color: XlsxWriter::Format::COLOR_YELLOW)
        series.set_fill(color: XlsxWriter::Format::COLOR_RED)
        cs.chart = chart
      end
    end
  end
end
