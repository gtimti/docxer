module Docxer
  module Word
    class WebSettings

      attr_accessor :options
      def initialize(options)
        @options = options
      end

      def render(zip)
        zip.put_next_entry('word/webSettings.xml')
        zip.write(Docxer.to_xml(document))
      end

      private
      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.webSettings('xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'mc:Ignorable' => "w14") do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }
            xml.optimizeForBrowser
            xml.allowPNG
          end
        end
      end


    end
  end
end