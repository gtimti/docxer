require 'docxer/word/medias/media'

module Docxer
  module Word
    class Medias

      attr_accessor :medias, :counter
      def initialize(options={})
        @medias = []
        @counter = 0
      end

      def add(media)
        @medias << media
        media
      end

      def sequence
        @counter += 1
      end

      def render(zip)
        @medias.each do |media|
          zip.put_next_entry("word/#{media.target}")
          zip.write(media.file.read)
        end
      end

    end
  end
end