# frozen_string_literal: true

require 'mkmf'
require 'English'

libxlsxwriter_dir = File.expand_path('libxlsxwriter', __dir__)
make_pid = spawn({ 'CFLAGS' => '-fPIC -O2' }, $make, chdir: libxlsxwriter_dir)
Process.wait(make_pid)
raise 'Make failed for xlsxwriter' unless $CHILD_STATUS.success?

$CFLAGS = "-I'#{libxlsxwriter_dir}/include' -g -Wall -O2"
$LDFLAGS = "-lz #{libxlsxwriter_dir}/lib/libxlsxwriter.a"

create_makefile 'xlsxwriter/xlsxwriter'
