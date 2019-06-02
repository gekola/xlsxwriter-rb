# frozen_string_literal: true

require_relative './xlsx-func-testcase'
require_relative './support/chart_test'

class TestChartAxis < XlsxWriterTestCase
  extend ChartTest

  DATA = [
    [1,  2,  3],
    [2,  4,  6],
    [3,  6,  9],
    [4,  8, 12],
    [5, 10, 15],
  ]

  DATA2 = [
    [1, 8,  3],
    [2, 7,  6],
    [3, 6,  9],
    [4, 8,  12],
    [5, 10, 15],
  ]

  chart_test 'chart_axis01', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 58_955_648
    chart.axis_id_2 = 68_446_848

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.name = 'XXX'
    chart.y_axis.name = 'YYY'
  end

  chart_test 'chart_axis02', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 43_704_320
    chart.axis_id_2 = 43_706_624

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.name = 'XXX'
    chart.y_axis.name = 'YYY'
  end

  chart_test 'chart_axis04', XlsxWriter::Workbook::Chart::SCATTER do |chart|
    chart.axis_id_1 = 46_891_776
    chart.axis_id_2 = 46_893_312

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

    chart.x_axis.name = 'XXX'
    chart.y_axis.name = 'YYY'
  end

  chart_test 'chart_axis05', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 47_076_480
    chart.axis_id_2 = 47_078_016

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.name = 'XXX'
    chart.y_axis.name = 'YYY'
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

  chart_test 'chart_axis07', XlsxWriter::Workbook::Chart::AREA, DATA2 do |chart|
    chart.axis_id_1 = 43_321_216
    chart.axis_id_2 = 47_077_248

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

    chart.x_axis.name = 'XXX'
    chart.y_axis.name = 'YYY'
  end

  chart_test 'chart_axis08', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 49_042_560
    chart.axis_id_2 = 61_199_872

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

    chart.y_axis.reverse = true
  end

  chart_test 'chart_axis09', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 49_043_712
    chart.axis_id_2 = 60_236_160

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.reverse = true
  end

  chart_test 'chart_axis10', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 41_012_608
    chart.axis_id_2 = 55_821_440

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

    chart.x_axis.reverse = true
    chart.y_axis.reverse = true
  end

  chart_test 'chart_axis11', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 45_705_472
    chart.axis_id_2 = 54_518_528

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.min = 0
    chart.x_axis.max = 20
  end

  chart_test 'chart_axis12', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 45_705_088
    chart.axis_id_2 = 54_517_760

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.y_axis.min = 0
    chart.y_axis.max = 16
  end

  chart_test 'chart_axis13', XlsxWriter::Workbook::Chart::SCATTER do |chart|
    chart.axis_id_1 = 54_045_312
    chart.axis_id_2 = 54_043_776

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

    chart.x_axis.min = 0
    chart.x_axis.max = 6
    chart.y_axis.min = 0
    chart.y_axis.max = 16
  end

  chart_test 'chart_axis15', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 45_705_856
    chart.axis_id_2 = 54_518_528

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.y_axis.major_unit = 2
    chart.y_axis.minor_unit = 0.4
  end

  chart_test 'chart_axis17', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 43_812_736
    chart.axis_id_2 = 45_705_088

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.y_axis.log_base = 10
  end

  chart_test 'chart_axis18', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 43_813_504
    chart.axis_id_2 = 45_705_472

    series = chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    series.invert_if_negative = true
  end

  chart_test 'chart_axis19', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 43_813_504
    chart.axis_id_2 = 45_705_472

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.label_position = XlsxWriter::Workbook::Chart::Axis::LABEL_POSITION_HIGH
    chart.y_axis.label_position = XlsxWriter::Workbook::Chart::Axis::LABEL_POSITION_LOW
  end

  chart_test 'chart_axis20', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 43_572_224
    chart.axis_id_2 = 43_812_352

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.label_position = XlsxWriter::Workbook::Chart::Axis::LABEL_POSITION_NEXT_TO
    chart.y_axis.label_position = XlsxWriter::Workbook::Chart::Axis::LABEL_POSITION_NONE
  end

  chart_test 'chart_axis21', XlsxWriter::Workbook::Chart::SCATTER do |chart|
    chart.axis_id_1 = 58_921_344
    chart.axis_id_2 = 54_519_680

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

    chart.x_axis.reverse = true
  end

  chart_test 'chart_axis22', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 86_799_104
    chart.axis_id_2 = 86_801_792

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.num_format = '#,##0.00'
  end

  chart_test 'chart_axis23', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 46_332_160
    chart.axis_id_2 = 47_470_848

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.num_format = 'dd/mm/yyyy'
    chart.y_axis.num_format = '0.00%'
  end

  chart_test 'chart_axis24', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 47_471_232
    chart.axis_id_2 = 48_509_696

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.num_format = 'dd/mm/yyyy'
    chart.y_axis.num_format = '0.00%'
    chart.x_axis.source_linked = 1
  end

  chart_test 'chart_axis25', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 47_471_232
    chart.axis_id_2 = 48_509_696

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.num_format = '[$¥-411]#,##0.00'
    chart.y_axis.num_format = '0.00%'
  end

  chart_test 'chart_axis26', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 73_048_448
    chart.axis_id_2 = 73_049_984

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.set_num_font rotation: 45, baseline: -1
  end

  chart_test 'chart_axis27', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 73_048_448
    chart.axis_id_2 = 73_049_984

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.set_num_font rotation: -35, baseline: -1
  end

  chart_test 'chart_axis28', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 45_451_904
    chart.axis_id_2 = 47_401_600

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.set_num_font rotation: 90, baseline: -1
  end

  chart_test 'chart_axis29', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 45_444_480
    chart.axis_id_2 = 47_402_368

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.set_num_font rotation: -90, baseline: -1
  end

  chart_test 'chart_axis30', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 69_200_896
    chart.axis_id_2 = 69_215_360

    chart.x_axis.position = XlsxWriter::Workbook::Chart::Axis::POSITION_ON_TICK

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'
  end

  chart_test 'chart_axis31', XlsxWriter::Workbook::Chart::BAR do |chart|
    chart.axis_id_1 = 90_616_960
    chart.axis_id_2 = 90_618_496

    chart.y_axis.position = XlsxWriter::Workbook::Chart::Axis::POSITION_ON_TICK

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'
  end

  chart_test 'chart_axis32', XlsxWriter::Workbook::Chart::AREA do |chart|
    chart.axis_id_1 = 96_171_520
    chart.axis_id_2 = 96_173_056

    chart.x_axis.position = XlsxWriter::Workbook::Chart::Axis::POSITION_BETWEEN

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'
  end

  chart_test 'chart_axis33', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 68_827_008
    chart.axis_id_2 = 68_898_816

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.name = 'XXX'
    chart.x_axis.set_name_font rotation: -45, baseline: -1

    chart.y_axis.name = 'YYY'
    chart.y_axis.set_name_font rotation: -45, baseline: -1
  end

  chart_test 'chart_axis34', XlsxWriter::Workbook::Chart::LINE do |chart|
    chart.axis_id_1 = 54_712_192
    chart.axis_id_2 = 54_713_728

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.interval_unit = 2
  end

  chart_test 'chart_axis35', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 63_008_128
    chart.axis_id_2 = 62_522_496

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.y_axis.set_line none: true
  end

  chart_test 'chart_axis36', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 45_501_056
    chart.axis_id_2 = 47_505_792

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.set_line none: true
  end

  chart_test 'chart_axis37', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 46_032_384
    chart.axis_id_2 = 48_088_960

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.set_line  color: XlsxWriter::Format::COLOR_YELLOW
    chart.y_axis.set_line  color: XlsxWriter::Format::COLOR_RED
  end

  chart_test 'chart_axis38', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 45_642_496
    chart.axis_id_2 = 45_644_416

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.y_axis.set_line color: XlsxWriter::Format::COLOR_YELLOW
    chart.y_axis.set_fill color: XlsxWriter::Format::COLOR_RED
  end

  chart_test 'chart_axis39', XlsxWriter::Workbook::Chart::SCATTER do |chart|
    chart.axis_id_1 = 45_884_928
    chart.axis_id_2 = 45_883_392

    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$A$1:$A$5', '=Sheet1!$C$1:$C$5'

    chart.x_axis.set_line none: true
    chart.y_axis.set_line none: true
  end

  chart_test 'chart_axis40', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 108_329_216
    chart.axis_id_2 = 108_635_264

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.interval_unit = 3
    chart.x_axis.interval_tick = 2
  end

  chart_test 'chart_axis41', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 108_321_024
    chart.axis_id_2 = 108_328_448

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.major_tick_mark = XlsxWriter::Workbook::Chart::Axis::TICK_MARK_NONE
    chart.x_axis.minor_tick_mark = XlsxWriter::Workbook::Chart::Axis::TICK_MARK_INSIDE
    chart.y_axis.minor_tick_mark = XlsxWriter::Workbook::Chart::Axis::TICK_MARK_CROSSING
  end

  chart_test 'chart_axis42', XlsxWriter::Workbook::Chart::COLUMN do |chart|
    chart.axis_id_1 = 61_296_640
    chart.axis_id_2 = 61_298_560

    chart.add_series '=Sheet1!$A$1:$A$5'
    chart.add_series '=Sheet1!$B$1:$B$5'
    chart.add_series '=Sheet1!$C$1:$C$5'

    chart.x_axis.label_align = XlsxWriter::Workbook::Chart::Axis::LABEL_ALIGN_RIGHT
  end
end
