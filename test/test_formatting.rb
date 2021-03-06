# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestFormatting < XlsxWriterTestCase
  test 'format01' do |wb|
    ws1 = wb.add_worksheet
    wb.add_worksheet 'Data Sheet'
    ws3 = wb.add_worksheet

    wb.add_format :unused, {}
    wb.add_format :format, bold: true
    wb.add_format :unused2, {}
    wb.add_format :unused3, {}

    ws1.write_string(0, 0, 'Foo', nil)
    ws1.write_number(1, 0, 123, nil)

    ws3.write_string(1, 1, 'Foo', nil)
    ws3.write_string(2, 1, 'Bar', :format)
    ws3.write_number(3, 2, 234, nil)
  end

  test 'format02' do |wb|
    ws = wb.add_worksheet

    ws.set_row(0, height: 30)
    wb.add_format(:format1, font_name: 'Arial',
                            bold: true,
                            align: XlsxWriter::Format::ALIGN_LEFT,
                            vertical_align: XlsxWriter::Format::ALIGN_VERTICAL_BOTTOM)
    wb.add_format(:format2, font_name: 'Arial',
                            bold: true,
                            rotation: 90,
                            align: XlsxWriter::Format::ALIGN_CENTER,
                            vertical_align: XlsxWriter::Format::ALIGN_VERTICAL_BOTTOM)

    ws.write_string(0, 0, 'Foo', :format1)
    ws.write_string(0, 1, 'Bar', :format2)
  end

  test 'format06' do |wb|
    ws = wb.add_worksheet

    wb.add_format :f1, num_format_index: 2
    wb.add_format :f2, num_format_index: 12

    ws.write_number(0, 0, 1.2222, nil)
    ws.write_number(1, 0, 1.2222, :f1)
    ws.write_number(2, 0, 1.2222, :f2)
    ws.write_number(3, 0, 1.2222, nil)
    ws.write_number(4, 0, 1.2222, nil)
  end

  test 'format07' do |wb|
    ws = wb.add_worksheet

    wb.add_format :f1, num_format: '0.000'
    wb.add_format :f2, num_format: '0.00000'
    wb.add_format :f3, num_format: '0.000000'

    ws.write_number(0, 0, 1.2222, nil)
    ws.write_number(1, 0, 1.2222, :f1)
    ws.write_number(2, 0, 1.2222, :f2)
    ws.write_number(3, 0, 1.2222, :f3)
    ws.write_number(4, 0, 1.2222, nil)
  end

  test 'format08' do |wb|
    ws = wb.add_worksheet

    wb.add_format :border1, bottom: XlsxWriter::Format::BORDER_THIN,
                            bottom_color: XlsxWriter::Format::COLOR_RED
    wb.add_format :border2, top: XlsxWriter::Format::BORDER_THIN,
                            top_color: XlsxWriter::Format::COLOR_RED
    wb.add_format :border3, left: XlsxWriter::Format::BORDER_THIN,
                            left_color: XlsxWriter::Format::COLOR_RED
    wb.add_format :border4, right: XlsxWriter::Format::BORDER_THIN,
                            right_color: XlsxWriter::Format::COLOR_RED
    wb.add_format :border5, border: XlsxWriter::Format::BORDER_THIN,
                            border_color: XlsxWriter::Format::COLOR_RED

    ws.write_blank(1, 1, :border1)
    ws.write_blank(3, 1, :border2)
    ws.write_blank(5, 1, :border3)
    ws.write_blank(7, 1, :border4)
    ws.write_blank(9, 1, :border5)
  end

  test 'format09' do |wb|
    ws = wb.add_worksheet

    wb.add_format :border1, border: XlsxWriter::Format::BORDER_HAIR,
                            border_color: XlsxWriter::Format::COLOR_RED
    wb.add_format :border2, diag_type: XlsxWriter::Format::DIAGONAL_BORDER_UP,
                            diag_color: XlsxWriter::Format::COLOR_RED
    wb.add_format :border3, diag_type: XlsxWriter::Format::DIAGONAL_BORDER_DOWN,
                            diag_color: XlsxWriter::Format::COLOR_RED
    wb.add_format :border4, diag_type: XlsxWriter::Format::DIAGONAL_BORDER_UP_DOWN,
                            diag_color: XlsxWriter::Format::COLOR_RED

    ws.write_blank(1, 1, :border1)
    ws.write_blank(3, 1, :border2)
    ws.write_blank(5, 1, :border3)
    ws.write_blank(7, 1, :border4)
  end

  test 'format10' do |wb|
    ws = wb.add_worksheet

    wb.add_format :border1, bg_color: XlsxWriter::Format::COLOR_RED
    wb.add_format :border2, bg_color: XlsxWriter::Format::COLOR_YELLOW,
                            pattern: XlsxWriter::Format::PATTERN_DARK_VERTICAL
    wb.add_format :border3, bg_color: XlsxWriter::Format::COLOR_YELLOW,
                            fg_color: XlsxWriter::Format::COLOR_RED,
                            pattern: XlsxWriter::Format::PATTERN_GRAY_0625

    ws.write_blank(1, 1, :border1)
    ws.write_blank(3, 1, :border2)
    ws.write_blank(5, 1, :border3)
  end

  test 'format12' do |wb|
    ws = wb.add_worksheet

    wb.add_format :top_left_bottom, bottom: XlsxWriter::Format::BORDER_THIN,
                                    left: XlsxWriter::Format::BORDER_THIN,
                                    top: XlsxWriter::Format::BORDER_THIN
    wb.add_format :top_bottom,      bottom: XlsxWriter::Format::BORDER_THIN,
                                    top: XlsxWriter::Format::BORDER_THIN
    wb.add_format :top_left,        left: XlsxWriter::Format::BORDER_THIN,
                                    top: XlsxWriter::Format::BORDER_THIN
    wb.add_format :unused,          left: XlsxWriter::Format::BORDER_THIN

    ws.write_string(1, 'B', 'test', :top_left_bottom)
    ws.write_string(1, 'D', 'test', :top_left)
    ws.write_string(1, 'F', 'test', :top_bottom)
  end

  test 'format15' do |wb|
    wb
      .add_format(:format1, bold: true)
      .add_format(:format2, bold: true, num_format_index: 1)
      .add_worksheet
      .write_number('A1', 1, :format1)
      .write_number('A2', 2, :format2)
  end

  test 'format50' do |wb|
    wb
      .add_format(:format1, num_format: '#,##0.00000')
      .add_format(:format2, num_format: '#,##0.0')
      .add_worksheet
      .set_column('A', 'A', width: 12)
      .write_number('A1', 1234.5, :format1)
      .write_number('A2', 1234.5, :format2)
  end

  test 'format51' do |wb|
    wb.add_format(:numf1, num_format: '0.0')
      .add_format(:numf3, num_format: '0.000')
      .add_format(:numf4, num_format: '0.0000')
      .add_format(:numf5, num_format: '0.00000')
      .add_worksheet
      .set_column(0, 0, width: 12)
      .write_number(0, 0, 123.456, :numf1)
      .write_number(1, 0, 123.456, :numf3)
      .write_number(2, 0, 123.456, :numf4)
      .write_number(3, 0, 123.456, :numf5)
  end

  test 'format52' do |wb|
    wb.add_format(:numf1, num_format: '0.0')
      .add_format(:numf3, num_format: '0.000')
      .add_format(:numf4, num_format: '0.0000')
      .add_format(:numf5, num_format: '0.00000')
      .add_format(:numf1_bold, num_format: '0.0', bold: true)
      .add_format(:numf3_bold, num_format: '0.000', bold: true)
      .add_format(:numf4_bold, num_format: '0.0000', bold: true)
      .add_format(:numf5_bold, num_format: '0.00000', bold: true)
      .add_worksheet
      .set_column(0, 0, width: 12)
      .write_number(0, 0, 123.456, :numf1)
      .write_number(1, 0, 123.456, :numf3)
      .write_number(2, 0, 123.456, :numf4)
      .write_number(3, 0, 123.456, :numf5)
      .write_number(4, 0, 123.456, :numf1_bold)
      .write_number(5, 0, 123.456, :numf3_bold)
      .write_number(6, 0, 123.456, :numf4_bold)
      .write_number(7, 0, 123.456, :numf5_bold)
  end
end
