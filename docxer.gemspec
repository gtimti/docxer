Gem::Specification.new do |s|
  s.name        = 'docxer'
  s.version     = '0.7.1'
  s.date        = '2012-03-20'
  s.summary     = 'docx files generator'
  s.description = 'Provides ability to create docx documents using Ruby'
  s.authors     = ["Timur Gabdrakipov"]
  s.email       = 'gtimti@gmail.com'
  s.homepage    = 'http://github.com/gtimti/docxer'
  s.license     = 'MIT'

  s.files = Dir["lib/**/*"] + ["LICENSE", "Rakefile", "README.md"]

  s.add_dependency 'nokogiri', '>= 1.6.3', '< 1.7'
  s.add_dependency 'rubyzip', '~> 0.9.9'
end
