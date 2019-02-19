#!/usr/bin/env ruby
# frozen_string_literal: true

base_dir = File.expand_path(File.join(File.dirname(__FILE__), ".."))
lib_dir = File.join(base_dir, "lib")
test_dir = File.join(base_dir, "test")

$LOAD_PATH.unshift(lib_dir)

require 'test/unit'
require 'xlsxwriter'

puts "XlsxWriter version #{XlsxWriter::VERSION} with libxlsxwriter #{XlsxWriter::LIBRARY_VERSION}"

exit Test::Unit::AutoRunner.run(true, test_dir)
