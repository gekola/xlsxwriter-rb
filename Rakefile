require 'rake/extensiontask'

task default: :test

spec = Gem::Specification.load('xlsxwriter.gemspec')
Rake::ExtensionTask.new('xlsxwriter', spec) do |ext|
  ext.lib_dir = 'lib/xlsxwriter'
end
task compile: :patch_dep

Gem::PackageTask.new(spec) do |pkg|
end

DEP_DIR='ext/xlsxwriter/libxlsxwriter'

desc "Checkout xlsxwriter C library"
task :patch_dep do
 `cd #{DEP_DIR} && git reset --hard`

  Dir['./dep_patches/*.patch'].each do |patch|
    `(cd #{DEP_DIR} && patch -N -p1) <#{patch}`
  end
end

desc 'Run specs'
task test: :compile do
  ruby('test/run-test.rb')
end
