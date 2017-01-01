require_relative './xlsx-func-testcase'

class TestPanes < XlsxWriterTestCase
  test 'panes01' do |wb|
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .freeze_panes(1, 'A')
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .freeze_panes(2, 'A')
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .freeze_panes(0, 'B')
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .freeze_panes(0, 'C')
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .freeze_panes(1, 'B')
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .freeze_panes(3, 'G')
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .freeze_panes(3, 6, 3, 6, 1)
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .split_panes(15, 0)
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .split_panes(30, 0)
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .split_panes(0, 8.46)
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .split_panes(0, 17.57)
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .split_panes(15, 8.46)
    wb.add_worksheet
      .write_string(0, 'A', 'Foo', nil)
      .split_panes(45, 54.14)
  end
end
