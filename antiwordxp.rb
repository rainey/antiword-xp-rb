#!/usr/bin/ruby
#Matt Smith/Shawn Rainey
#antiword-xp.rb - Convert docx files to plaintext

#  add an each_wrapped_line method to the String class
class String
	
	#Takes a width to wrap, defaulted to $wrapWidth, and a paragraph separator
	# the separator is inserted after each paragraph
	# along with an option to add the seperator after single-line "paragraphs"
	# which fit on one line with no wrapping.  This will likely always be false
	def each_wrapped_line(cols = $wrapWidth, p_seperator="\n", seperateSingle=false)
		
		lines = []
		self.each_line { |line|
						
			words = line.split
			wrapped_line = ""
			
			seperate = seperateSingle
	
			words.each { | word | 
				word.strip!
				#Is there room for the next word? 
				if (wrapped_line.length + word.length) <= cols || cols == 0
					wrapped_line << word
					#Add Space if it will fit
					wrapped_line << " " unless wrapped_line.length == cols
				else
					#Always use seperator when paragraphs span more than one line
					seperate = true
					
					lines << wrapped_line
			
					#If the word length is bigger than the number of columns
					# add it to lines. Otherwise add it to the next wrapped line
					if word.length + 1 > cols
						lines << word
					else
						wrapped_line = word + " "
					end
				end
			}
			wrapped_line += p_seperator if seperate
			lines << wrapped_line
		}
		
		#Yield lines if block given, otherwise return them.
		lines.each { |line| yield(line) } if block_given?
		return lines
	end
end

#Found out the hard way that the env. var $COLUMNS is not exported...
#So we do this instead
begin
	IO.popen("tput cols"){ |process| $consoleWidth = process.read.to_i }
	$wrapWidth = $consoleWidth
rescue Errno::ENOENT
	$wrapWidth = $consoleWidth = 80
end

def usage
	"Usage: #{$0} takes a .doc or .docx formatted word document. It can be called either by piping the document to antiword, or by calling `#{$0} filename`
	".each_wrapped_line($consoleWidth) { | line | puts line }

puts "
Arguments:"

"-w## or -w ##	Set wrap with.  If not specified, uses console width or 80 if console width cannot be determined.

--notimeout Disable input timeout.  This could be necessary for large files or files from external sources.  Only needed when piping in a word file, and not when one is specified in the programs argument list.".each_wrapped_line($consoleWidth) { | line | puts line }


puts"
Examples: 
	$#{$0} < mydoc.doc[x]
	$#{$0} mydoc.doc[x] -w 60 --notimeout
	$cat mydoc.doc[x] | #{$0} -w80
"
end

stdinTimeout = 5
filename = nil

#Generate a hash string from a random number to seperate the arguments
#See antiword.rb.txt
require 'digest'
arg_sep = "<"

#Choose a new argument seperator until sep. not found in arguments
# use the TR because we don't want digits in the seperator
until ARGV.join("") !~ /#{arg_sep}/ 
	arg_sep = Digest::hexencode(Digest::SHA2.new().digest(rand().to_s)).tr("0-9", "G-P")
end

argstring = ARGV.join(arg_sep)

#if we can find a -h or -help, or can't find a good indicator
#of a doc/docx
if (argstring =~ /(?:#{arg_sep}|^)-+h(?:elp)?(?=#{arg_sep}|$)/)
	usage
	exit(1)
else
	temp_fname = nil
	
	argtokens = { 
				"--notimeout" => \
					lambda { | matchData | stdinTimeout = 0 if matchData
				 	},
	 			
	 			#Take a width value.  can be "-w10" or "-w 10" on the command line  
	 			"-w(?:#{arg_sep})?(\\d+)" => \
 					lambda { | matchData | 
 						$wrapWidth = matchData.to_a[1].to_i unless matchData == nil
 					},
	 			
	 			"(.+\\.docx?)" => \
 					lambda { | matchData | 
 						temp_fname = matchData.to_a[1] unless matchData == nil
 					}
	 		} 
	
	argtokens.each_pair { | expression, callback |
		#Call the callback function with the matchdata from the expression-injected RE
		# expression is matched between arg_sep or start/end anchors 
		# ending seperator is not included in match, but beginning seperator is.
		callback.call(argstring.match(/.*(?:#{arg_sep}|^)(?>#{expression})(?=#{arg_sep}|$)/))
	}
	
	
	
	#If a file name is given, 
	# test if the given filename exists
	if temp_fname != nil
		if File.exist?(temp_fname)
			filename = temp_fname

		else
			puts "#{temp_fname} does not exist!"
			usage
			Process.exit(1);
		end
	end
	
	#Clear the argument list of known arguments
	argtokens.each_key { | expression |
		argstring.gsub!(/(?:#{arg_sep}|^)#{expression}(?=#{arg_sep}|$)/, "")
	}
	
	#If there are still arguments left, and the file name has not been
	# assigned, assume that the unrecognized arg is meant to be the file,
	# and output the error message + usage.
	#This allows us to ignore garbage arguments when a file is supplied.
	# Unfortunately, they will still be problematic when the word file is
	# piped in.
	if(temp_fname == nil && !argstring.empty?)
		argstring.gsub!(/#{arg_sep}/, "")
		puts "#{argstring} is not a valid word file."
		usage
		Process.exit(1)
	end
end	

process_xml = true;
#Copy contents of stdin to antiword.zip
if filename == nil
	begin
		require 'timeout'
		Timeout::timeout(stdinTimeout) do
			File.open("antiword_temp.zip", "w") { |file| file.write($stdin.read) }
			filename = "antiword_temp.zip"
		end
	rescue Timeout::Error
		File.delete("antiword_temp.zip")
		
		"Timed out.  This can happen if you piped in a very large file, or if you did not specify a file at all.  To remedy this with very large files, add the --notimeout argument when calling #{$0}.

".each_wrapped_line { |line| puts line }
		
		usage
		Process.exit(1)
	end
end

document = String.new

gotContents = true

#Set to "if false" if RubyZip is causing problems
# and it will use the system's unzip instead.
if true
	require 'rubygems'
	require 'zip/zipfilesystem'
	begin
		Zip::ZipFile.open(filename) { | awContents |
			document = awContents.read("word/document.xml")
		}
	rescue Zip::ZipError
		gotContents = false;
	end
else
	#unzip options: pipe output to stdout, only extract word/document.xml
	#result.read captures stdout from the opened process.
	IO.popen("unzip -p antiword_temp.zip word/document.xml 2> /dev/null") { |result| document = result.read }
	gotContents = ($? == 0)
end


#If the unzip failed
unless gotContents
	process_xml = nil
	
	#If the filename isn't antiword_temp, and this is a doc, copy
	#the file to antiword_temp.  Do this to avoid having to escape the filename
	unless filename == "antiword_temp.zip"
		File.open("antiword_temp.zip", "w") { |awfile| 
			File.open(filename) { | inFile | awfile.write(inFile.read) }
		filename = "antiword_temp.zip" 
		# ^^^ antiword_temp.zip doesn't get deleted unless filename is set to this 
	}
	end
	
	#Try to process with system's antiword, maybe it's an old doc file.
	#Set antiword's options: one paragraph per line, text mode, no images.
	#	This matches the format of document.xml with the tags processed
	IO.popen('antiword antiword_temp.zip -w 0 -t -i 1 2> /dev/null') {
	 |result| document = result.read 
	}
		
	#if antiword failed
	#You're SOL
	unless $? == 0
		$stderr.write("Unsupported format\n")
		usage
		File.delete("antiword_temp.zip")
		Process.exit 1
	end
end

if(process_xml)
	replacements = []
	#Remove line breaks.  There are none in MS-Words's XML, but
	#Could change in the future.  Or could have been generated
	#using something else
	replacements << [ /\n|\r/, '']
	
	#Add seperators where column tags are using pipe, unless last in row
	replacements << [ /<\/w:p><\/w:tc>(?!<\/w:tr>)/, " | " ]
	
	#list elements, may add more soon
	replacements << [ /<w:numPr>/, "-" ]
	
	#Tabbed Columns
	replacements << [ /<w:tab[^\/]*\/>/, " " ]
	
	#Substitute end paragraph tag with newline
	#Effectively, this should treat each paragraph on one line
	replacements << [ /<\/w:p>/, "\r\n" ]
	
	#insert [pic] to replace graphics.
	replacements << [ /<pic:pic[^>]*>/, '[pic]']
	
	replacements << [ /<wp:posOffset>\d+?<\/wp:posOffset>/, "" ] 
	
	#Remove all other tags
	replacements << [ /<[^>]*>/, "" ]

	#Not sure if any other replacements need to be made, but this should 
	# make it easy enough to add more
	replacements << [ /&lt;/ , '<' ] <<
					[ /&gt;/,  '>' ] << 
					[ /&amp;/, "&"] <<
					[ /&quot;/, '"'] <<
					[ /&apos;/, "'" ]
					
	replacements.each { | replacement | 
		document.gsub!(replacement[0], replacement[1]) 
	}
	
	#Some UTF-8 characters don't print
	#This translates from utf-8 to ascii
	require "iconv"
	document = Iconv.conv("ascii//translit", "UTF-8", document)	
end

begin
	document.each_wrapped_line {|line| $stdout.write( line + "\n") }
rescue Errno::EPIPE
end

File.delete("antiword_temp.zip") if filename == "antiword_temp.zip"
