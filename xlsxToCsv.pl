#!/usr/bin/perl

use strict;
use warnings;
use Spreadsheet::XLSX;

my $no_Sheets = 0;
my $no_Sheet_rows = 0;
my $no_rows = 0;
my $no_nhm_cells = 0;
my $no_cells = 0;
my $no_original_cells = 0;
#my $prefix = '';
 
#my $excel = Spreadsheet::XLSX -> new ('test.xlsx');
my $excel = Spreadsheet::XLSX -> new ('test2.xlsx');
#my $excel = Spreadsheet::XLSX -> new ('CRM365_SB_Schema.xlsx');
#my $excel = Spreadsheet::XLSX -> new ('CRM2013_SB_Schema.xlsx');
my $line;
foreach my $sheet (@{$excel -> {Worksheet}}) {
    #printf("Sheet: %s\n", $sheet->{Name});

    $sheet -> {MaxRow} ||= $sheet -> {MinRow};
	$no_Sheet_rows = 0;
    foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
			if (($no_Sheets == 1) && ($no_Sheet_rows == 0)) {
			$line = '"'."Table".'"'.",";;
			}
			else
			{		$line = '"'.$sheet->{Name}.'"'.",";
			}
        $sheet -> {MaxCol} ||= $sheet -> {MinCol};
		my $counted = '';
        foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
		    $no_original_cells++;
			
            my $cell = $sheet -> {Cells} [$row] [$col];
            if ($cell) {
                $line .= "\"".$cell -> {Val}."\",";
            
				my $prefix = $cell -> {Val};
				$prefix = substr($prefix, 0, 3);
				if (($prefix eq  "sb_") && ($counted ne  "Y"))  {
					$no_cells++;
					$counted = "Y";
				}
				$prefix = substr($cell -> {Val}, 0, 6);
				if (($prefix eq  "sbnhm_") && ($counted ne  "Y"))  {
					$no_nhm_cells++;
					$counted = "Y";
				}
			}
        }
		if ($no_Sheets == 0) {
			if ($no_Sheet_rows == 0) {
				print "Sheet not required.\n"
				}
			}
			else
			{
			if (($no_Sheets == 1) && ($no_Sheet_rows == 0)) {
					chomp($line);    # Heading line
					$line =~ s/,$//; 
					print "$line\n";
				}
				elsif ($no_Sheet_rows ne 0) {
					chomp($line);
					$line =~ s/,$//; 
					print "$line\n";
					}
				
			}


		$line = '';
		$no_rows++;
		$no_Sheet_rows++;
    }
	$no_Sheets++;
}

##  Post table create change.
##  update [scratchpad_colif_dev].[dbo].[CRM2013]  set TableName = replace(stuff(Entity,1,charindex('(',Entity),''),')','') ;
##

printf("\n\nNumber of sheets : $no_Sheets \n");
printf("Number of rows : $no_rows \n");
printf("Number of rows with cells with sb_ : $no_cells \n");
printf("Number of rows with cells with sbnhm_ : $no_nhm_cells \n");
printf("Total Number of all original cells : $no_original_cells \n");