# coding: utf-8
# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestRichString < XlsxWriterTestCase
  test('rich_string01') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', ['a', ['bc', :bold], 'defg'])
  end

  test('rich_string02') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', ['abcd', ['ef', :italic], 'g'])
  end

  test('rich_string03') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', [['abc', :bold], 'defg'])
  end

  test('rich_string04') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', [['abc', :bold], ['defg', :italic]])
  end

  test('rich_string05') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .set_column('A', 'A', width: 30)
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3',
        XlsxWriter::RichString.new(wb) <<
          'This is ' << ['bold', :bold] << ' and this is ' << ['italic', :italic]
      )
  end


  test('rich_string06') do |wb|
    wb
      .add_format(:red, font_color: XlsxWriter::Format::COLOR_RED)
      .add_worksheet
      .write_string('A1', 'Foo', :red)
      .write_string('A2', 'Bar')
      .write_rich_string('A3', ['ab', ['cde', :red], 'fg'])
  end

  test('rich_string07') do |wb|
    t1, t2, t3, t4 = [
      ['a', ['bc', :bold], 'defg'],
      ['a', ['bcdef', :bold], 'g'],
      ['abc', ['de', :italic], 'fg'],
      [['abcd', :italic], ['efg', nil]],
    ].map { |parts| XlsxWriter::RichString.new(wb, parts) }
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', t1)
      .write_rich_string('B4', t3)
      .write_rich_string('C5', t1)
      .write_rich_string('D6', t3)
      .write_rich_string('E7', t2)
      .write_rich_string('F8', t4)
  end

  test('rich_string08') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_format(:centered, align: XlsxWriter::Format::ALIGN_CENTER)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', ['ab', ['cd', :bold], 'efg'], :centered)
  end

  test('rich_string09') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
    ws = wb.add_worksheet
    ws
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', ['a', ['bc', :bold], 'defg'])

    assert_raise(XlsxWriter::Error, 'Function parameter validation error.') do
      ws.write_rich_string('A3', ['', ['bc', :bold], 'defg'])
    end
    assert_raise(XlsxWriter::Error, 'Function parameter validation error.') do
      ws.write_rich_string('A3', [])
    end
    assert_raise(XlsxWriter::Error, 'Function parameter validation error.') do
      ws.write_rich_string('A3', [['foo', :bold]])
    end
  end

  test('rich_string10') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', [' a', ['bc', :bold], 'defg '])
  end

  test('rich_string11') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', ['a', ['☺', :bold], 'defg'])
  end

  test('rich_string12') do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_format(:wrap, text_wrap: true)
      .add_worksheet
      .set_column('A', 'A', width: 30)
      .set_row(2, height: 60)
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', XlsxWriter::RichString.new(wb) <<
          "This is\n" << ["bold\n", :bold] << "and this is\n" << ['italic', :italic],
        :wrap
      )
  end
end
