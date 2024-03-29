# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestImage < XlsxWriterTestCase
  test 'image01' do |wb|
    wb.add_worksheet.insert_image(8, 'E', image_path('red.png'), {})
  end

  test 'image02' do |wb|
    wb.add_worksheet.insert_image(6, 'D', image_path('yellow.png'), x_offset: 1, y_offset: 2)
  end

  test 'image03' do |wb|
    wb.add_worksheet.insert_image(8, 'E', image_path('red.jpg'), {})
  end

  test 'image04' do |wb|
    wb.add_worksheet.insert_image(8, 'E', image_path('red.bmp'), {})
  end

  test 'image05' do |wb|
    ws = wb.add_worksheet
    ws.insert_image(0, 'A', image_path('blue.png'), {})
    ws.insert_image(2, 'B', image_path('red.jpg'), {})
    ws.insert_image(4, 'D', image_path('yellow.jpg'), {})
    ws.insert_image(8, 'F', image_path('grey.png'), {})
  end

  test 'image07' do |wb|
    wb.add_worksheet.insert_image(8, 'E', image_path('red.png'), {})
    wb.add_worksheet.insert_image(8, 'E', image_path('yellow.png'), {})
  end

  test 'image08' do |wb|
    wb.add_worksheet.insert_image(2, 'B', image_path('grey.png'), scale: 0.5)
  end

  test 'image09' do |wb|
    wb.add_worksheet.insert_image(8, 'E', image_path('red_64x20.png'), {})
  end

  test 'image10' do |wb|
    wb.add_worksheet.insert_image(1, 'C', image_path('logo.png'), {})
  end

  test 'image11' do |wb|
    wb.add_worksheet.insert_image(1, 'C', image_path('logo.png'), x_offset: 8, y_offset: 5)
  end

  test 'image12' do |wb|
    ws = wb.add_worksheet
    ws.set_row(1, height: 75)
    ws.set_column(2, 2, width: 32)
    ws.insert_image(1, 'C', image_path('logo.png'), {})
  end

  test 'image13' do |wb|
    ws = wb.add_worksheet
    ws.set_row(1, height: 75)
    ws.set_column(2, 2, width: 32)
    ws.insert_image(1, 'C', image_path('logo.png'), x_offset: 8, y_offset: 5)
  end

  test 'image14' do |wb|
    ws = wb.add_worksheet
    ws.set_row(1, height: 4.5)
    ws.set_row(2, height: 35.25)
    ws.set_column(2, 4, width: 3.29)
    ws.set_column(5, 5, width: 10.71)
    ws.insert_image(1, 'C', image_path('logo.png'), {})
  end

  test 'image15' do |wb|
    ws = wb.add_worksheet
    ws.set_row(1, height: 4.5)
    ws.set_row(2, height: 35.25)
    ws.set_column(2, 4, width: 3.29)
    ws.set_column(5, 5, width: 10.71)
    ws.insert_image(1, 'C', image_path('logo.png'), x_offset: 13, y_offset: 2)
  end

  test 'image29' do |wb|
    wb.add_worksheet.insert_image(0, 10, image_path('red_208.png'), x_offset: -210, y_offset: 1)
  end

  test 'image36' do |wb|
    wb.add_worksheet.insert_image('E9', image_path('happy.jpg'))
  end

  test 'image44' do |wb|
    wb.add_worksheet
      .insert_image('E9', image_path('red.png'))
      .set_row(8, height: 30)
  end

  test 'image45' do |wb|
    wb.add_worksheet
      .insert_image('E9', image_path('red.png'), object_position: XlsxWriter::Worksheet::OBJECT_MOVE_AND_SIZE_AFTER)
      .set_row(8, height: 30, hide: true)
  end

  test 'image46' do |wb|
    wb.add_worksheet
      .insert_image('E9', image_path('red.png'), y_offset: 4, object_position: XlsxWriter::Worksheet::OBJECT_MOVE_AND_SIZE_AFTER)
      .set_row(8, height: 30, hide: true)
  end

  test 'image47' do |wb|
    wb.add_worksheet
      .insert_image('E9', image_path('red.png'))
      .set_row(9, hide: true)
  end

  test('image81', ref_file_name: 'image01') do |wb|
    wb.add_worksheet
      .insert_image_buffer('E9', File.binread(image_path('red.png')), description: 'red.png')
  end

  test 'image82' do |wb|
    wb.add_worksheet
      .insert_image_buffer('E9', File.binread(image_path('red.png')))
  end

  test('image83', ref_file_name: 'image02') do |wb|
    wb.add_worksheet
      .insert_image_buffer('D7', File.binread(image_path('yellow.png')), x_offset: 1,
                                                                         y_offset: 2,
                                                                         description: 'yellow.png')
  end
end
