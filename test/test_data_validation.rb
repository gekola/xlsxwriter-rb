# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestDataValidation < XlsxWriterTestCase
  test 'data_validation01' do |wb|
    wb.add_worksheet
      .add_data_validation 'C2', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                 value: %w[Foo Bar Baz]
  end

  test 'data_validation02' do |wb|
    wb.add_worksheet
      .add_data_validation 1, 2, validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                 value: %w[Foo Bar Baz],
                                 input_title: 'This is the input title',
                                 input_message: 'This is the input message'
  end

  test 'data_validation03' do |wb|
    wb.add_worksheet
      .add_data_validation('C2', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                 value: %w[Foo Bar Baz],
                                 input_title: 'This is the input title',
                                 input_message: 'This is the input message')
      .add_data_validation('D6', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                 value: %w[Foobar Foobas Foobat Foobau Foobav Foobaw Foobax
                                           Foobay Foobaz Foobba Foobbb Foobbc Foobbd Foobbe
                                           Foobbf Foobbg Foobbh Foobbi Foobbj Foobbk Foobbl
                                           Foobbm Foobbn Foobbo Foobbp Foobbq Foobbr Foobbs
                                           Foobbt Foobbu Foobbv Foobbw Foobbx Foobby Foobbz
                                           Foobca End],
                                 input_title: 'This is the longest input title1',
                                 input_message: "This is the longest input message #{'a' * 221}")
  end

  test 'data_validation04', ref_file_name: 'data_validation02' do |wb|
    ws = wb.add_worksheet
    ws.add_data_validation('C2', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                 value: %w[Foo Bar Baz],
                                 input_title: 'This is the input title',
                                 input_message: 'This is the input message')
    assert_raise(XlsxWriter::Error.new('Parameter exceeds Excel\'s limit of 32 characters.')) do
      ws.add_data_validation('D6', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                   value: %w[Foobar Foobas Foobat Foobau Foobav Foobaw Foobax
                                             Foobay Foobaz Foobba Foobbb Foobbc Foobbd Foobbe
                                             Foobbf Foobbg Foobbh Foobbi Foobbj Foobbk Foobbl
                                             Foobbm Foobbn Foobbo Foobbp Foobbq Foobbr Foobbs
                                             Foobbt Foobbu Foobbv Foobbw Foobbx Foobby Foobbz
                                             Foobca End],
                                   input_title: 'This is the longest input title12',
                                   input_message: "This is the longest input message #{'a' * 221}")
    end
  end

  test 'data_validation05', ref_file_name: 'data_validation02' do |wb|
    ws = wb.add_worksheet
    ws.add_data_validation('C2', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                 value: %w[Foo Bar Baz],
                                 input_title: 'This is the input title',
                                 input_message: 'This is the input message')
    assert_raise(XlsxWriter::Error.new('Parameter exceeds Excel\'s limit of 255 characters.')) do
      ws.add_data_validation('D6', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                   value: %w[Foobar Foobas Foobat Foobau Foobav Foobaw Foobax
                                             Foobay Foobaz Foobba Foobbb Foobbc Foobbd Foobbe
                                             Foobbf Foobbg Foobbh Foobbi Foobbj Foobbk Foobbl
                                             Foobbm Foobbn Foobbo Foobbp Foobbq Foobbr Foobbs
                                             Foobbt Foobbu Foobbv Foobbw Foobbx Foobby Foobbz
                                             Foobca End],
                                   input_title: 'This is the longest input title1',
                                   input_message: "This is the longest input message #{'a' * 222}")
    end
  end

  test 'data_validation06', ref_file_name: 'data_validation02' do |wb|
    ws = wb.add_worksheet
    ws.add_data_validation('C2', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                 value: %w[Foo Bar Baz],
                                 input_title: 'This is the input title',
                                 input_message: 'This is the input message')
    assert_raise(XlsxWriter::Error.new('Parameter exceeds Excel\'s limit of 255 characters.')) do
      ws.add_data_validation('D6', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                   value: %w[Foobar Foobas Foobat Foobau Foobav Foobaw Foobax
                                             Foobay Foobaz Foobba Foobbb Foobbc Foobbd Foobbe
                                             Foobbf Foobbg Foobbh Foobbi Foobbj Foobbk Foobbl
                                             Foobbm Foobbn Foobbo Foobbp Foobbq Foobbr Foobbs
                                             Foobbt Foobbu Foobbv Foobbw Foobbx Foobby Foobbz
                                             Foobca End1],
                                   input_title: 'This is the longest input title1',
                                   input_message: "This is the longest input message #{'a' * 221}")
    end
  end

  test 'data_validation07' do |wb|
    wb.add_worksheet
      .add_data_validation 'C2', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_LIST,
                                 value: %w[coffee café]
  end

  test 'data_validation08' do |wb|
    wb.add_worksheet
      .add_data_validation 'C2:C2', validate: XlsxWriter::Worksheet::VALIDATION_TYPE_ANY,
                                    input_title: 'This is the input title',
                                    input_message: 'This is the input message'
  end
end
