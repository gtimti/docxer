module Docxer
  module Word
    class Document

      include Helpers

      attr_accessor :options, :fonts, :setings, :styles, :numbering, :effects, :web_settings, :themes, :relationships, :footnotes, :endnotes, :footers, :headers

      def initialize(options)
        @options = options
        @content = []
      end

      def add_media(file, options={})
        media = Medias::Media.new(medias.sequence, file, options)
        rel = relationships.add(media.content_type, media.target)
        media.set_sequence(rel.id)
        medias.add(media)
      end

      def add_footer(footer, options={})
        footer.id = footers.sequence
        rel = relationships.add(footer.content_type, footer.target)
        footer.sequence = rel.id
        footers.add(footer)
        @footer = footer
      end

      def add_header(header, options={})
        header.id = headers.sequence
        rel = relationships.add(header.content_type, header.target)
        header.sequence = rel.id
        headers.add(header)
        @header = header
      end

      def fonts
        @fonts ||= Fonts.new
      end

      def settings
        @settings ||= Settings.new
      end

      def styles
        @styles ||= Styles.new
      end

      def numbering
        @numbering ||= Numbering.new
      end

      def effects
        @effects ||= Effects.new
      end

      def web_settings
        @web_settings ||= WebSettings.new({})
      end

      def themes
        @themes ||= Themes.new
      end

      def medias
        @medias ||= Medias.new
      end

      def relationships
        @relationships ||= Relationships.new
      end

      def footnotes
        @footnotes ||= Footnotes.new
      end

      def endnotes
        @endnotes ||= Endnotes.new
      end

      def footers
        @footers ||= Footers.new
      end

      def headers
        @headers ||= Headers.new
      end

      def render(zip)
        zip.put_next_entry('word/document.xml')
        zip.write(Docxer.to_xml(document))

        fonts.render(zip)
        settings.render(zip)
        styles.render(zip)
        effects.render(zip)
        numbering.render(zip)
        web_settings.render(zip)
        themes.render(zip)
        footnotes.render(zip)
        endnotes.render(zip)
        headers.render(zip)
        footers.render(zip)
        medias.render(zip)
        relationships.render(zip)
      end

      private
      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.document('xmlns:wpc' => "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", 'xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:o' => "urn:schemas-microsoft-com:office:office", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:m' => "http://schemas.openxmlformats.org/officeDocument/2006/math", 'xmlns:v' => "urn:schemas-microsoft-com:vml", 'xmlns:wp14' => "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", 'xmlns:wp' => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", 'xmlns:w10' => "urn:schemas-microsoft-com:office:word", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'xmlns:wpg' => "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup", 'xmlns:wpi' => "http://schemas.microsoft.com/office/word/2010/wordprocessingInk", 'xmlns:wne' => "http://schemas.microsoft.com/office/word/2006/wordml", 'xmlns:wps' => "http://schemas.microsoft.com/office/word/2010/wordprocessingShape", 'mc:Ignorable' => "w14 wp14") do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }

            xml['w'].body do

              content.each do |element|
                element.render(xml)
              end

              xml['w'].p do
                xml['w'].bookmarkStart( 'w:id' => "0", 'w:name' => "_GoBack")
                xml['w'].bookmarkEnd('w:id' => "0")
              end
              xml['w'].sectPr do
                xml['w'].footerReference( 'w:type' => "default", 'r:id' => @footer.sequence ) if @footer
                xml['w'].headerReference( 'w:type' => "default", 'r:id' => @header.sequence ) if @header
                xml['w'].titlePg
                xml['w'].pgSz( 'w:w' => 12240, 'w:h' => 15840 )
                xml['w'].pgMar( 'w:top' => 800, 'w:right' => 940, 'w:bottom' => 1000, 'w:left' => 940, 'w:header' => 190, 'w:footer' => 200, 'w:gutter' => 0)
                xml['w'].cols( 'w:space' => 720)
                xml['w'].docGrid( 'w:linePitch' => "360")
              end
            end
          end
        end
      end

      def content
        @content
      end

    end
  end
end
