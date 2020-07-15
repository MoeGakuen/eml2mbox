#!/usr/bin/ruby
# encoding: utf-8

require "date"
require 'fileutils'

#=======================================================#
# Class that encapsulates the processing file in memory #
#=======================================================#
class FileInMemory
    
    ZoneOffset = {
        # Standard zones by RFC 2822
        'UTC' => '0000', 
        'UT' => '0000', 'GMT' => '0000',
        'EST' => '-0500', 'EDT' => '-0400',
        'CST' => '-0600', 'CDT' => '-0500',
        'MST' => '-0700', 'MDT' => '-0600',
        'PST' => '-0800', 'PDT' => '-0700',
    }   
    
    def initialize()
        @lines = Array.new
        @counter = 1          # keep the 0 position for the From_ line
        @from = nil           # from part of the From_ line
        @prefrom = nil        # buffer for multiline From:
        @date = nil           # date part of the From_ line
    end

    def addLine(line)
        line = line.force_encoding("ISO-8859-1").encode("utf-8", replace: nil)

        # If the line is a 'false' From line, add a '>' to its beggining
        line = line.sub(/From/, '>From') if line =~ /^From/ and @from!=nil

        # If previous line was a two-liner From header without address concatenate both
        if @prefrom != nil
            line = @prefrom + " " + line
            @prefrom = nil
        end
        
        # If the line is the first valid From line, save it (without the line break)
        if line =~ /^From:\s.*/ and @from==nil
            if line =~ /.*@/
                @from = line.sub(/From:/,'From')
                @from = @from.chop    # Remove line break(s)
                @from = standardizeFrom(@from)
            end
        end

        if line =~ /^Date:\s/ and @date==nil
            # Parse content of the Date header and convert to the mbox standard for the From_ line
            @date = line.sub(/Date:\s/,'')
            year, month, day, hour, minute, second, timezone, wday = DateTime._parse(@date, false).values_at(:year, :mon, :mday, :hour, :min, :sec, :zone, :wday)
            # Need to convert the timezone from a string to a 4 digit offset
            unless timezone =~ /[+|-]\d*/
                timezone=ZoneOffset[timezone]
            end
            begin
                time = Time.gm(year,month,day,hour,minute,second)
                @date = formMboxDate(time,timezone)
            rescue
                @date = nil
                $errors = true
                print "[skipping bad date]"
            end
        end

        # Now add the line to the array
        line = fixLineEndings(line)
        @lines[@counter]=line
        @counter+=1
    end

    # Forms the first line (from + date) and returns all the lines
    # Returns all the lines in the file
    def getProcessedLines()
        if @from != nil
            # Add from and date to the first line
            if @date==nil
                $errors = true
                print "[replacing bad date with now]"
                @date=formMboxDate(Time.now,nil)
            end
            @lines[0] = @from + " " + @date 
            
            @lines[0] = fixLineEndings(@lines[0])
            @lines[@counter] = ""
            return @lines
        end
        # else don't return anything
    end

    # Fixes CR/LFs
    def fixLineEndings(line)
        return line
    end
end

#================#
# Helper methods #
#================#

# Converts: 'From "some one <aa@aa.aa>" <aa@aa.aa>' -> 'From aa@aa.aa'
def standardizeFrom(fromLine)
    # Get indexes of last "<" and ">" in line
    openIndex = fromLine.rindex('<')
    closeIndex = fromLine.rindex('>')
    if openIndex!=nil and closeIndex!=nil
        fromLine = fromLine[0..4]+fromLine[openIndex+1..closeIndex-1]
    end
    # else leave as it is - it is either already well formed or is invalid
    return fromLine
end

# Returns a mbox postmark formatted date.
# If timezone is unknown, it is skipped.
# mbox date format used is described here:
# http://www.broobles.com/eml2mbox/mbox.html
def formMboxDate(time,timezone)
    if timezone==nil
        return time.strftime("%a %b %d %H:%M:%S %Y")
    else
        return time.strftime("%a %b %d %H:%M:%S %Y "+timezone.to_s)
    end
end

# 检查并创建多级目录
def checkDir(path)
    if not File.directory?(path)
        FileUtils.mkdir_p(path)
    end
end

# 核心处理函数
def handle(searchPath)
    emlFiles = Dir.glob("#{searchPath}*.{eml,mai}", File::FNM_CASEFOLD)
    if emlFiles.size == 0
        msg = "[#{searchPath.gsub('/', '\\').chop()}]"
        print msg, "目录没有 EML 文件".rjust(110 - msg.length), "\n"
        return
    end

    mboxPath = savePath + "#{searchPath.gsub('/', '\\').chop()}.mbox"
    puts "========== [#{searchPath.gsub('/', '\\').chop()}]（#{emlFiles.size}）==========\n\n"
    puts "输出路径：" + mboxPath + "\n\n"

    if File.exist?(mboxPath)
        print "文件已存在！请选择：[A]追加  [O]覆盖  [C]跳过（默认）："
        sel = STDIN.gets.chomp
        if sel == 'A' or sel == 'a'
            fileHandle = File.new(mboxPath, "a");
            puts
        elsif sel == 'O' or sel == 'o'
            fileHandle = File.new(mboxPath, "w");
            puts
        else
            puts "\n\n\n"
            return
        end
    else
        tempPath = searchPath.gsub('/', '\\').chop()
        tempIndex = tempPath.rindex('\\')
        tempIndex = tempIndex != nil ? tempIndex : 0
        checkDir(savePath + tempPath[0, tempIndex])
        fileHandle = File.new(mboxPath, "w");
    end    

    fileNum = 0
    errorNum = 0
    emlFiles.each() do |i|
        isError = false
        fileNum += 1
        fileNumStr = fileNum.to_s.rjust("#{emlFiles.size}".length)
        msg = "#{fileNumStr}/#{emlFiles.size}：#{i}"
        print msg
        memoryFile = FileInMemory.new()
        File.open(i).each {|item| memoryFile.addLine(item)}

        if not ARGV[1]
            lines = memoryFile.getProcessedLines
            if lines == nil
                isError = true
                print '[跳过没有常规发件人的邮件]'.ljust(105 - msg.length)
            else
                lines.each {|line| fileHandle.puts line}
            end
        end

        if isError
            print "\n"
            errorNum += 1
        else
            print "\r"
        end
    end
    fileHandle.close

    puts "处理完毕，有 #{errorNum} 个文件出错".ljust(120)
    puts "\n\n\n"
end


#===============#
#     Main      #
#===============#

$stdout.sync = true

system 'title EML To Mbox'
system 'cls'

workPath = File.dirname(__FILE__).gsub('/', '\\') + '\\'
savePath = workPath + 'mbox\\'

if ARGV[0] != nil
    emlDir = ARGV[0]
    if emlDir.rindex(':/') == nil and emlDir.rindex(':\\') == nil
        emlDir = workPath + emlDir
    end
else
    emlDir = workPath
end

if File.directory?(emlDir)
    Dir.chdir(emlDir)
    #puts "\n扫描路径：[" + emlDir + "]\n\n"
    puts
else
    puts "\n[#{emlDir}] 不是一个目录，或目录不存在，请指定有效的目录！\n\n\n"
    exit(0)
end

handle(emlDir)
Dir.glob('**/').each do |searchPath|
    handle(searchPath)
end
