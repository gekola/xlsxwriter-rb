# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestOptimize < XlsxWriterTestCase
  test('optimize01', constant_memory: true) do |wb|
    wb.add_worksheet
      .write_string(0, 0, 'Hello', nil)
      .write_string(0, 0, 'Hello', nil)
      .write_number(1, 0, 123, nil)
  end

  test('optimize02', constant_memory: true) do |wb|
    wb.add_worksheet
      .write_string(0, 0, 'Hello', nil)
      .write_number(1, 0, 123, nil)
      .write_string(0, 'G', 'World', nil)
    # G1 should be ignored since a later row has already been written.
  end

  test('optimize04', constant_memory: true) do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', ['a', ['bc', :bold], 'defg'])
  end

  test('optimize05', constant_memory: true) do |wb|
    t1, t2, t3, t4 = [
      ['a', ['bc', :bold], 'defg'],
      ['a', ['bcdef', :bold], 'g'],
      ['abc', ['de', :italic], 'fg'],
      [['abcd', :italic], ['efg', nil]]
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

  test('optimize06', constant_memory: true) do |wb|
    ws = wb.add_worksheet
    ws.write_string(0, 0, '_x0000_', nil)
    (1..127).each { |i| ws.write_string(i, 0, i.chr, nil) unless i == 34 }
  end

  test('optimize08', constant_memory: true) do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_format(:italic, italic: true)
      .add_worksheet
      .write_string('A1', 'Foo', :bold)
      .write_string('A2', 'Bar', :italic)
      .write_rich_string('A3', [' a', ['bc', :bold], 'defg '])
  end

  test('optimize21', constant_memory: true) do |wb|
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .write_string(2, 'C', ' Foo', nil)
      .write_string(4, 'E', 'Foo ', nil)
      .write_string(6, 'A', "\tFoo\t", nil)
  end

  test('optimize22', constant_memory: true) do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_worksheet
      .set_column(0, 0, width: 36, format: :bold)
  end

  test('optimize23', constant_memory: true) do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_worksheet
      .set_row(0, height: 20, format: :bold)
  end

  test('optimize24', constant_memory: true) do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_worksheet
      .set_row(0, height: 20, format: :bold)
      .write_string(0, 0, 'Foo')
  end

  test('optimize25', constant_memory: true) do |wb|
    wb
      .add_format(:bold, bold: true)
      .add_worksheet
      .set_row(0, height: 20, format: :bold)
      .write_string(2, 0, 'Foo')
  end

  test('optimize26', constant_memory: true) do |wb|
    wb.add_worksheet.write_string(2, 2, 'café')
  end
end
