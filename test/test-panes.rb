# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestPanes < XlsxWriterTestCase
  test 'panes01' do |wb|
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .freeze_panes(1, 'A')
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .freeze_panes('A3')
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .freeze_panes('B1')
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .freeze_panes('C1')
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .freeze_panes('B2')
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .freeze_panes('G4')
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .freeze_panes('G4', 'G4', 1)
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .split_panes(15, 0)
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .split_panes(30, 0)
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .split_panes(0, 8.46)
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .split_panes(0, 17.57)
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .split_panes(15, 8.46)
    wb.add_worksheet
      .write_string('A1', 'Foo')
      .split_panes(45, 54.14)
  end
end
