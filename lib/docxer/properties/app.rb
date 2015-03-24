module Docxer
  class Properties
    class App

      attr_accessor :options

      def initialize(options)
        @options = options
      end

      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.Properties('xmlns' => "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", 'xmlns:vt' => "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes") do
            xml.Template "Normal.dotm"
            xml.TotalTime 0
            xml.Pages 1
            xml.Words 0
            xml.Characters 0
            xml.Application "Docxer Ruby Gem"
            xml.DocSecurity 0
            xml.Lines 0
            xml.Paragraphs 0
            xml.ScaleCrop false
            xml.Company options[:creator]
            xml.LinksUpToDate false
            xml.CharactersWithSpaces 0
            xml.SharedDoc false
            xml.HyperlinksChanged false
            xml.AppVersion "14.0000"
          end
        end
      end


    end
  end
end
