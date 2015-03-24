module Docxer
  class Properties
    class Core
      
      attr_accessor :options

      def initialize(options)
        @options = options
      end
      
      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.coreProperties('xmlns:cp' => 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', 'xmlns:dc' => 'http://purl.org/dc/elements/1.1/', 'xmlns:dcterms' => 'http://purl.org/dc/terms/', 'xmlns:dcmitype' => 'http://purl.org/dc/dcmitype/', 'xmlns:xsi' => 'http://www.w3.org/2001/XMLSchema-instance') do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix=="cp" }
            # xml['dc'].title options[:title]
            xml['dc'].creator do
              options[:creator]
            end
            xml['cp'].lastModifiedBy do
              options[:creator]
            end
            xml['cp'].revision 1
            xml['dcterms'].created(Time.now.strftime('%Y-%m-%dT%H:%M:%SZ'), 'xsi:type' => 'dcterms:W3CDTF')
            xml['dcterms'].modified(Time.now.strftime('%Y-%m-%dT%H:%M:%SZ'), 'xsi:type' => 'dcterms:W3CDTF')
          end
        end
      end
      

    end
  end
end