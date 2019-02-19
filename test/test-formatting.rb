# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestFormatting < XlsxWriterTestCase
  test 'format01' do |wb|
    ws1 = wb.add_worksheet
    ws2 = wb.add_worksheet 'Data Sheet'
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
                            align:  XlsxWriter::Format::ALIGN_CENTER,
                            vertical_align:  XlsxWriter::Format::ALIGN_VERTICAL_BOTTOM)

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
                                    left:   XlsxWriter::Format::BORDER_THIN,
                                    top:    XlsxWriter::Format::BORDER_THIN
    wb.add_format :top_bottom,      bottom: XlsxWriter::Format::BORDER_THIN,
                                    top:    XlsxWriter::Format::BORDER_THIN
    wb.add_format :top_left,        left:   XlsxWriter::Format::BORDER_THIN,
                                    top:    XlsxWriter::Format::BORDER_THIN
    wb.add_format :unused,          left:   XlsxWriter::Format::BORDER_THIN

    ws.write_string(1, 'B', 'test', :top_left_bottom)
    ws.write_string(1, 'D', 'test', :top_left)
    ws.write_string(1, 'F', 'test', :top_bottom)
  end
end
