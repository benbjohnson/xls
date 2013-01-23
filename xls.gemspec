# -*- encoding: utf-8 -*-
$:.unshift File.expand_path("../lib", __FILE__)

require 'xls/version'

Gem::Specification.new do |s|
  s.name        = "xls"
  s.version     = Xls::VERSION
  s.authors     = ["Ben Johnson"]
  s.email       = ["benbjohnson@yahoo.com"]
  s.homepage    = "http://github.com/benbjohnson/xls"
  s.summary     = "A command line utilty for working with data in Excel."
  s.executables = ['xls']

  s.add_dependency('commander', '~> 4.1.3')
  s.add_dependency('ruby-progressbar', '~> 1.0.2')
  s.add_dependency('spreadsheet', '~> 0.7.6')

  s.add_development_dependency('rake', '~> 0.9.2.2')
  s.add_development_dependency('minitest', '~> 3.5.0')
  s.add_development_dependency('mocha', '~> 0.12.5')
  s.add_dependency('unindentable', '~> 0.1.0')

  s.test_files   = Dir.glob("test/**/*")
  s.files        = Dir.glob("lib/**/*") + %w(README.md)
  s.require_path = 'lib'
end
