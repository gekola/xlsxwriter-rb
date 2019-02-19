# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestProtect < XlsxWriterTestCase
  test 'protect02' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    wb.add_format :unlocked, unlocked: true
    wb.add_format :hidden, unlocked: true, hidden: true
    wb.add_worksheet do |ws|
      ws.protect
      ws.write_number(0, 'A', 1, nil)
      ws.write_number(1, 'A', 2, :unlocked)
      ws.write_number(2, 'A', 3, :hidden)
    end
  end

  test 'protect03' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    wb.add_format :unlocked, unlocked: true
    wb.add_format :hidden, unlocked: true, hidden: true
    wb.add_worksheet do |ws|
      ws.protect 'password'
      ws.write_number(0, 'A', 1, nil)
      ws.write_number(1, 'A', 2, :unlocked)
      ws.write_number(2, 'A', 3, :hidden)
    end
  end
end
