module Docxer
  module Word
    class Styles

      attr_accessor :options
      def initialize(options={})
        @options = options
      end

      def render(zip)
        zip.put_next_entry('word/styles.xml')
        zip.write(Docxer.to_xml(document))
      end

      private
      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.styles('xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'mc:Ignorable' => "w14") do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }
            xml['w'].docDefaults do
              xml['w'].rPrDefault do
                xml['w'].rPr do
                  xml['w'].rFonts('w:asciiTheme' => "minorHAnsi", 'w:eastAsiaTheme' => "minorHAnsi", 'w:hAnsiTheme' => "minorHAnsi", 'w:cstheme' => "minorBidi")
                  xml['w'].sz('w:val' => 22)
                  xml['w'].szCs('w:val' => 22)
                  xml['w'].lang('w:val' => "en-US", 'w:eastAsia' => "en-US", 'w:bidi' => "ar-SA")
                end
              end
              xml['w'].pPrDefault do
                xml['w'].pPr do
                  xml['w'].spacing('w:after' => 200, 'w:line' => 276, 'w:lineRule' => "auto")
                end
              end
            end
            xml['w'].latentStyles( 'w:defLockedState' => 0, 'w:defUIPriority' => 99, 'w:defSemiHidden' => 1, 'w:defUnhideWhenUsed' => 1, 'w:defQFormat' => 0, 'w:count' => 267) do
              xml['w'].lsdException( 'w:name' => "Normal", 'w:semiHidden' => 0, 'w:uiPriority' => 0, 'w:unhideWhenUsed' => 0, 'w:qFormat' => 1)
              xml['w'].lsdException( 'w:name' => "heading 1", 'w:semiHidden' => 0, 'w:uiPriority' => 9, 'w:unhideWhenUsed' => 0, 'w:qFormat' => 1)
              xml['w'].lsdException( 'w:name' => "heading 2", 'w:uiPriority' => "9", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "heading 3", 'w:uiPriority' => "9", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "heading 4", 'w:uiPriority' => "9", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "heading 5", 'w:uiPriority' => "9", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "heading 6", 'w:uiPriority' => "9", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "heading 7", 'w:uiPriority' => "9", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "heading 8", 'w:uiPriority' => "9", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "heading 9", 'w:uiPriority' => "9", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "toc 1", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "toc 2", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "toc 3", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "toc 4", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "toc 5", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "toc 6", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "toc 7", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "toc 8", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "toc 9", 'w:uiPriority' => "39")
              xml['w'].lsdException( 'w:name' => "caption", 'w:uiPriority' => "35", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Title", 'w:semiHidden' => "0", 'w:uiPriority' => "10", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Default Paragraph Font", 'w:uiPriority' => "1")
              xml['w'].lsdException( 'w:name' => "Subtitle", 'w:semiHidden' => "0", 'w:uiPriority' => "11", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Strong", 'w:semiHidden' => "0", 'w:uiPriority' => "22", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Emphasis", 'w:semiHidden' => "0", 'w:uiPriority' => "20", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Table Grid", 'w:semiHidden' => "0", 'w:uiPriority' => "59", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Placeholder Text", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "No Spacing", 'w:semiHidden' => "0", 'w:uiPriority' => "1", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Light Shading", 'w:semiHidden' => "0", 'w:uiPriority' => "60", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light List", 'w:semiHidden' => "0", 'w:uiPriority' => "61", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Grid", 'w:semiHidden' => "0", 'w:uiPriority' => "62", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 1", 'w:semiHidden' => "0", 'w:uiPriority' => "63", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 2", 'w:semiHidden' => "0", 'w:uiPriority' => "64", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 1", 'w:semiHidden' => "0", 'w:uiPriority' => "65", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 2", 'w:semiHidden' => "0", 'w:uiPriority' => "66", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 1", 'w:semiHidden' => "0", 'w:uiPriority' => "67", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 2", 'w:semiHidden' => "0", 'w:uiPriority' => "68", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 3", 'w:semiHidden' => "0", 'w:uiPriority' => "69", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Dark List", 'w:semiHidden' => "0", 'w:uiPriority' => "70", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Shading", 'w:semiHidden' => "0", 'w:uiPriority' => "71", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful List", 'w:semiHidden' => "0", 'w:uiPriority' => "72", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Grid", 'w:semiHidden' => "0", 'w:uiPriority' => "73", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Shading Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "60", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light List Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "61", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Grid Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "62", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 1 Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "63", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 2 Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "64", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 1 Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "65", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Revision", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "List Paragraph", 'w:semiHidden' => "0", 'w:uiPriority' => "34", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Quote", 'w:semiHidden' => "0", 'w:uiPriority' => "29", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Intense Quote", 'w:semiHidden' => "0", 'w:uiPriority' => "30", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Medium List 2 Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "66", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 1 Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "67", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 2 Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "68", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 3 Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "69", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Dark List Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "70", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Shading Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "71", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful List Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "72", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Grid Accent 1", 'w:semiHidden' => "0", 'w:uiPriority' => "73", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Shading Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "60", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light List Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "61", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Grid Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "62", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 1 Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "63", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 2 Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "64", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 1 Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "65", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 2 Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "66", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 1 Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "67", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 2 Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "68", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 3 Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "69", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Dark List Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "70", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Shading Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "71", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful List Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "72", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Grid Accent 2", 'w:semiHidden' => "0", 'w:uiPriority' => "73", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Shading Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "60", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light List Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "61", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Grid Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "62", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 1 Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "63", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 2 Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "64", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 1 Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "65", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 2 Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "66", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 1 Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "67", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 2 Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "68", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 3 Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "69", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Dark List Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "70", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Shading Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "71", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful List Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "72", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Grid Accent 3", 'w:semiHidden' => "0", 'w:uiPriority' => "73", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Shading Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "60", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light List Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "61", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Grid Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "62", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 1 Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "63", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 2 Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "64", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 1 Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "65", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 2 Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "66", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 1 Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "67", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 2 Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "68", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 3 Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "69", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Dark List Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "70", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Shading Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "71", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful List Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "72", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Grid Accent 4", 'w:semiHidden' => "0", 'w:uiPriority' => "73", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Shading Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "60", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light List Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "61", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Grid Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "62", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 1 Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "63", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 2 Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "64", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 1 Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "65", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 2 Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "66", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 1 Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "67", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 2 Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "68", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 3 Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "69", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Dark List Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "70", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Shading Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "71", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful List Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "72", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Grid Accent 5", 'w:semiHidden' => "0", 'w:uiPriority' => "73", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Shading Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "60", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light List Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "61", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Light Grid Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "62", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 1 Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "63", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Shading 2 Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "64", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 1 Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "65", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium List 2 Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "66", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 1 Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "67", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 2 Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "68", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Medium Grid 3 Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "69", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Dark List Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "70", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Shading Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "71", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful List Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "72", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Colorful Grid Accent 6", 'w:semiHidden' => "0", 'w:uiPriority' => "73", 'w:unhideWhenUsed' => "0")
              xml['w'].lsdException( 'w:name' => "Subtle Emphasis", 'w:semiHidden' => "0", 'w:uiPriority' => "19", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Intense Emphasis", 'w:semiHidden' => "0", 'w:uiPriority' => "21", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Subtle Reference", 'w:semiHidden' => "0", 'w:uiPriority' => "31", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Intense Reference", 'w:semiHidden' => "0", 'w:uiPriority' => "32", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Book Title", 'w:semiHidden' => "0", 'w:uiPriority' => "33", 'w:unhideWhenUsed' => "0", 'w:qFormat' => "1")
              xml['w'].lsdException( 'w:name' => "Bibliography", 'w:uiPriority' => "37")
              xml['w'].lsdException( 'w:name' => "TOC Heading", 'w:uiPriority' => "39", 'w:qFormat' => "1")
            end
            xml['w'].style( 'w:type' => "paragraph", 'w:default' => 1, 'w:styleId' => "Normal" ) do
              xml['w'].name( 'w:val' => "Normal" )
              xml['w'].qFormat
            end
            xml['w'].style( 'w:type' => "character", 'w:default' => 1, 'w:styleId' => "DefaultParagraphFont" ) do
              xml['w'].name( 'w:val' => "Default Paragraph Font" )
              xml['w'].uiPriority( 'w:val' => '1' )
              xml['w'].semiHidden
              xml['w'].unhideWhenUsed
            end
            xml['w'].style( 'w:type' => "table", 'w:default' => 1, 'w:styleId' => "TableNormal") do
              xml['w'].name( 'w:val' => "Normal Table")
              xml['w'].uiPriority( 'w:val' => 99 )
              xml['w'].semiHidden
              xml['w'].unhideWhenUsed
              xml['w'].tblPr do
                xml['w'].tblInd( 'w:w' => 0, 'w:type' => "dxa" )
                xml['w'].tblCellMar do
                  xml['w'].top( 'w:w' => 0, 'w:type' => "dxa" )
                  xml['w'].left( 'w:w' => 108, 'w:type' => "dxa" )
                  xml['w'].bottom( 'w:w' => 0, 'w:type' => "dxa" )
                  xml['w'].right( 'w:w' => 108, 'w:type' => "dxa" )
                end
              end
            end
            xml['w'].style( 'w:type' => "table", 'w:styleId' => "TableCustom" ) do
              xml['w'].name( 'w:val' => "Table Grid" )
              xml['w'].basedOn( 'w:val' => "TableNormal" )
              xml['w'].uiPriority( 'w:val' => "59" )
              xml['w'].pPr do
                xml['w'].spacing( 'w:after' => "0", 'w:line' => "240", 'w:lineRule' => "auto" )
              end
              xml['w'].tblPr do
                xml['w'].tblInd( 'w:w' => 0, 'w:type' => "dxa" )
                xml['w'].tblBorders do
                  xml['w'].top( 'w:val' => "single", 'w:sz' => 4, 'w:space' => 0, 'w:color' => "auto" )
                  xml['w'].left( 'w:val' => "single", 'w:sz' => 4, 'w:space' => 0, 'w:color' => "auto" )
                  xml['w'].bottom( 'w:val' => "single", 'w:sz' => 4, 'w:space' => 0, 'w:color' => "auto" )
                  xml['w'].right( 'w:val' => "single", 'w:sz' => 4, 'w:space' => 0, 'w:color' => "auto" )
                  xml['w'].insideH( 'w:val' => "single", 'w:sz' => 4, 'w:space' => 0, 'w:color' => "auto" )
                  xml['w'].insideV( 'w:val' => "single", 'w:sz' => 4, 'w:space' => 0, 'w:color' => "auto" )
                end
                xml['w'].tblCellMar do
                  xml['w'].top( 'w:w' => 0, 'w:type' => "dxa" )
                  xml['w'].left( 'w:w' => 108, 'w:type' => "dxa" )
                  xml['w'].bottom( 'w:w' => 0, 'w:type' => "dxa" )
                  xml['w'].right( 'w:w' => 108, 'w:type' => "dxa" )
                end
              end
            end
            xml['w'].style( 'w:type' => "numbering", 'w:default' => 1, 'w:styleId' => "NoList") do
              xml['w'].name( 'w:val' => "No List" )
              xml['w'].uiPriority( 'w:val' => "99" )
              xml['w'].semiHidden
              xml['w'].unhideWhenUsed
            end
          end
        end
      end
    end
  end
end
