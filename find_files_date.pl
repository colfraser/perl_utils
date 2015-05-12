#!/usr/local/bin/perl
# File all files. Output modified date, path and filename
# perl find_files_date.pl | sort -r       - for reverse date output
use strict;
use File::Find qw(find);
use POSIX 'strftime';
#use File::stat;
my $dir     = 'C:\Perl\PM\NHM\test';
#my $dir     = 'Y:\CRM';
my $pattern = '';  # blank for everything
find sub {  
		open my $fh, $File::Find::name;
		my ($dev,$ino,$mode,$nlink,$uid,$gid,$rdev,$size,
              $atime,$mtime,$ctime,$blksize,$blocks)
                  = stat($fh);
		my $when = strftime '%Y %m %d', localtime $mtime;print $when . " ";
		print $File::Find::name . "\n" if /$pattern/;
}, $dir;

