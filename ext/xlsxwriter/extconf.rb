require 'mkmf'

libxlsxwriter_dir = File.expand_path('../libxlsxwriter', __FILE__)
make_pid = spawn({'CFLAGS' => '-fPIC'}, $make, chdir: libxlsxwriter_dir)
Process.wait(make_pid)
raise 'Make failed for xlsxwriter' unless $? == 0

# enable_config('static', true)
# find_library('xlsxwriter', 'workbook_new', File.expand_path('../libxlsxwriter/lib', __FILE__))
$CFLAGS="-I'#{libxlsxwriter_dir}/include' -g -Wall"
# $LDFLAGS="-L./libxlsxwriter/lib/ -Wl,-Bstatic -lxlsxwriter -Wl,-Bdynamic"
$LDFLAGS="-lz #{libxlsxwriter_dir}/lib/libxlsxwriter.a"

create_makefile 'xlsxwriter/xlsxwriter'
