#encoding: utf-8

module Docxer
  module Word
    class Footers
      class Footer

        attr_accessor :id, :sequence, :options, :content, :relationships
        def initialize(options={})
          @options = options
          @content = []
          @relationships = []

          if block_given?
            yield self
          else

          end
        end

        def content_type
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
        end

        def target
          "footer#{id}.xml"
        end

        def render(zip)
          zip.put_next_entry("word/#{target}")
          zip.write(Docxer.to_xml(document))

          if !@relationships.empty?
            zip.put_next_entry("word/_rels/#{target}.rels")
            zip.write(Docxer.to_xml(relationships))
          end
        end

        def text(text, options={})
          options = @options.merge(options)
          element = Docxer::Word::Contents::Paragraph.new(options) do |p|
            p.text(text)
          end
          @content << element
          element
        end

        def page_numbers
          numbers = PageNumbers.new
          @content << numbers
          numbers
        end

        class PageNumbers

          attr_accessor :options
          def initialize(options={})
            @options = options
          end

          def render(xml)
            xml['w'].sdt do
              xml['w'].sdtPr do
                xml['w'].id( 'w:val' => "-472213903" )
                xml['w'].docPartObj do
                  xml['w'].docPartGallery( 'w:val' => "Page Numbers (Bottom of Page)" )
                  xml['w'].docPartUnique
                end
              end
              xml['w'].sdtContent do
                xml['w'].p do
                  xml['w'].pPr do
                    xml['w'].jc( 'w:val' => @options[:align] || 'right' )
                  end
                  xml['w'].r do
                    xml['w'].fldChar( 'w:fldCharType' => "begin" )
                  end
                  xml['w'].r do
                    xml['w'].instrText "PAGE   \* MERGEFORMAT"
                  end
                  xml['w'].r do
                    xml['w'].fldChar( 'w:fldCharType' => "separate" )
                  end
                  xml['w'].r do
                    xml['w'].fldChar( 'w:fldCharType' => "end" )
                  end
                end
              end
            end
          end
        end

        private
        def document
          Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
            xml.ftr( 'xmlns:wpc' => "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", 'xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:o' => "urn:schemas-microsoft-com:office:office", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:m' => "http://schemas.openxmlformats.org/officeDocument/2006/math", 'xmlns:v' => "urn:schemas-microsoft-com:vml", 'xmlns:wp14' => "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", 'xmlns:wp' => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", 'xmlns:w10' => "urn:schemas-microsoft-com:office:word", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'xmlns:wpg' => "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup", 'xmlns:wpi' => "http://schemas.microsoft.com/office/word/2010/wordprocessingInk", 'xmlns:wne' => "http://schemas.microsoft.com/office/word/2006/wordml", 'xmlns:wps' => "http://schemas.microsoft.com/office/word/2010/wordprocessingShape", 'mc:Ignorable' => "w14 wp14" ) do
              xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }
              @content.each do |element|
                element.render(xml)
              end
            end
          end
        end

        def relationships
          Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
            xml.Relationships(xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships') do
              @relationships.each do |relationship|
                xml.Relationship('Id' => relationship.sequence, 'Type' => relationship.content_type, 'Target' => relationship.target)
              end
            end
          end
        end

      end
    end
  end
end
