module Roo
  module Utils
    extend self

    LETTERS = ('A'..'Z').to_a

    def split_coordinate(str)
      @split_coordinate ||= {}

      @split_coordinate[str] ||= begin
                                   letter, number = split_coord(str)
                                   x = letter_to_number(letter)
                                   y = number
                                   [y, x]
                                 end
    end

    alias_method :ref_to_key, :split_coordinate

    # copied from https://github.com/weshatheleopard/rubyXL
    # Converts +row+ and +col+ zero-based indices to Excel-style cell reference
    # (0) A...Z, AA...AZ, BA... ...ZZ, AAA... ...AZZ, BAA... ...XFD (16383)
    def ind2ref(row = 0, col = 0)
      str = ''

      loop do
        x = col % 26
        str = ('A'.ord + x).chr + str
        col = (col / 26).floor - 1
        break if col < 0
      end

      str += (row + 1).to_s
    end

    # copied from https://github.com/weshatheleopard/rubyXL
    # Converts Excel-style cell reference to +row+ and +col+ zero-based indices.
    def ref2ind(str)
      return [ -1, -1 ] unless str =~ /\A([A-Z]+)(\d+)\Z/

      col = 0
      $1.each_byte { |chr| col = col * 26 + (chr - 64) }
      [ $2.to_i - 1, col - 1 ]
    end

    def split_coord(s)
      if s =~ /([a-zA-Z]+)([0-9]+)/
        letter = Regexp.last_match[1]
        number = Regexp.last_match[2].to_i
      else
        fail ArgumentError
      end
      [letter, number]
    end

    # convert a number to something like 'AB' (1 => 'A', 2 => 'B', ...)
    def number_to_letter(num)
      result = ""

      until num.zero?
        num, index = (num - 1).divmod(26)
        result.prepend(LETTERS[index])
      end

      result
    end

    def letter_to_number(letters)
      @letter_to_number ||= {}
      @letter_to_number[letters] ||= begin
                                       result = 0

                                       # :bytes method returns an enumerator in 1.9.3 and an array in 2.0+
                                       letters.bytes.to_a.map{|b| b > 96 ? b - 96 : b - 64 }.reverse.each_with_index{ |num, i| result += num * 26 ** i }

                                       result
                                     end
    end

    # Compute upper bound for cells in a given cell range.
    def num_cells_in_range(str)
      cells = str.split(':')
      return 1 if cells.count == 1
      raise ArgumentError.new("invalid range string: #{str}. Supported range format 'A1:B2'") if cells.count != 2
      x1, y1 = split_coordinate(cells[0])
      x2, y2 = split_coordinate(cells[1])
      (x2 - (x1 - 1)) * (y2 - (y1 - 1))
    end

    def load_xml(path)
      ::File.open(path, 'rb') do |file|
        ::Nokogiri::XML(file)
      end
    end

    # Yield each element of a given type ('row', 'c', etc.) to caller
    def each_element(path, elements)
      Nokogiri::XML::Reader(::File.open(path, 'rb'), nil, nil, Nokogiri::XML::ParseOptions::NOBLANKS).each do |node|
        next unless node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT && Array(elements).include?(node.name)
        yield Nokogiri::XML(node.outer_xml).root if block_given?
      end
    end
  end
end
