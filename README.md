# XlsxWriter

[![Build Status](https://travis-ci.com/gekola/xlsxwriter-rb.svg?branch=master)](https://travis-ci.com/gekola/xlsxwriter-rb)
[![Gem Version](https://badge.fury.io/rb/xlsxwriter.svg)](https://badge.fury.io/rb/xlsxwriter)

## Description

Ruby binding to jmcnamara's [libxlswriter](https://github.com/jmcnamara/libxlsxwriter).

[Documentation link](http://www.rubydoc.info/gems/xlsxwriter/XlsxWriter).

### Usage

The following code snippet creates file `test.xlsx` in current directory with a header and one thousand row with random data.

```ruby
require 'xlsxwriter'

XlsxWriter::Workbook.new('test.xlsx') do |wb|
  wb.add_format(:numf, num_format_index: 4) # Excel standard fixed point format (two numbers after point)
  wb.add_format(:header, bg_color: 0XCCCCCC, bold: true, bottom: XlsxWriter::Format::BORDER_THIN)

  wb.add_worksheet('Worksheet 1') do |ws|
    ws.add_row(['Number', 'String'], style: :header)
    1_000.times do |i|
      ws.add_row([rand * 10_000, 'test string %0d' % (rand * 1000)], style: [:numf])
    end
  end
end
```

### Motivation

This gem has been written for generating large xlsx files (millions of rows). Pure ruby xlsx libraries do not perform very well in such use cases, consuming lots of memory for intermediate ruby objects and being generally slower (see the benchmark result below, benchmark code is located under `bench/simple_100k_rows.rb`).

    Rehearsal ----------------------------------------------
    xlsxwriter   0.601120   0.015773   0.616893 (  0.573616)
    axlsx       18.731412   2.845668  21.577080 ( 17.039327)
    ------------------------------------ total: 22.193973sec
    
                     user     system      total        real
    xlsxwriter   0.613193   0.000000   0.613193 (  0.557582)
    axlsx       19.055665   2.771983  21.827648 ( 17.259542)

Comparing this gem functionality to [axlsx](https://github.com/randym/axlsx) xlsxwriter lacks ability to read data from the current workbook state and does not have any integration with rails.
