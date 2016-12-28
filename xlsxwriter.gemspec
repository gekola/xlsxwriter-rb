Gem::Specification.new do |s|
  s.name = 'xlsxwriter'.freeze
  s.version = '0.0.1'.freeze
  s.summary = 'Ruby interface to libxlsxwriter'.freeze
  s.authors = ['Nick H'.freeze]
  s.files = ['Rakefile',
             'lib/xlsxwriter.rb',
             *Dir['ext/xlsxwriter/*.{c,h,rb}'],
             'ext/xlsxwriter/libxlsxwriter/Makefile',
             'ext/xlsxwriter/libxlsxwriter/lib/.gitignore',
             *Dir['ext/xlsxwriter/libxlsxwriter/{src,third_party}/**/Makefile'],
             *Dir['ext/xlsxwriter/libxlsxwriter/{src,third_party,include}/**/*.{c,h}']].map(&:freeze)
  s.extensions = Dir['ext/**/extconf.rb'].freeze

  s.add_development_dependency 'rake-compiler'
  s.add_development_dependency 'test-unit'
  s.add_development_dependency 'rubyzip'
end
