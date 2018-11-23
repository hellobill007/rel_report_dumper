use strict;
use warnings;
use Spreadsheet::Reader::ExcelXML;
use Data::Dumper;

my $rel_file=shift;
 
my $parser   = Spreadsheet::Reader::ExcelXML->new();
my $workbook = $parser->parse( $rel_file);

if ( !defined $workbook ) {
	die $parser->error(), "\n";
}
my $criterion=3;  #  percentage 
my $print_all_data=0;

my $sheet_index=-1;
my $test_item_array;
my $test_item_hash;
my $test_item_gap_hash;
my $rel_hrs_array;
#my $rel_hrs_gap_array;
for my $worksheet ( $workbook->worksheets() ) {
	$sheet_index++;
	#next if($sheet_index==0);
	my ( $row_min, $row_max ) = $worksheet->row_range();
	my ( $col_min, $col_max ) = $worksheet->col_range();
	$row_max=1000 unless($row_max);
	$col_max=1000 unless($col_max);
	push @$test_item_array,$worksheet->get_name;
	#print $worksheet->get_name,$/;
	$rel_hrs_array=();
	for my $col ( 1 .. $col_max ) {
		my $cell = $worksheet->get_cell( 1, $col );
		last if($cell->value() =~ /\%/);
		push @$rel_hrs_array,$cell->value();
	}
	#print "@$rel_hrs_array",$/;
	#next;
	my $index_hrs=0;
	for my $col ( 1 .. @$rel_hrs_array ) {
		for my $row ( 2 .. $row_max ) {
			my $cell = $worksheet->get_cell( $row, $col );
			last unless $cell;
			last if ($cell eq 'EOR' or $cell eq 'EOF');
			#last if($col==0 && not $cell->value()=~/\d+/);
			# print "Row, Col    = ($row, $col)\t";
			#print $cell->value(),"$/";
			push @{$test_item_hash->{$worksheet->get_name}[$index_hrs]},$cell->value() if($col);
			#$index_hrs++;
			#print "Unformatted = ", $cell->unformatted(), "\n";
		}
		$index_hrs++
		#print $/;
	}
	#last;# In order not to read all sheets
}
#print Dumper( $test_item_hash);
#exit;
foreach my $item (@$test_item_array){
	my @hrs_array=@{$test_item_hash->{$item}};
	my $index_hrs_gap=0;
	for my $index_hrs (1 .. @hrs_array-1){
		for my $index_count (0 .. @{$hrs_array[0]}-1){
			my $diff=$hrs_array[$index_hrs][$index_count]-$hrs_array[0][$index_count];
			if($diff==0){
				#$test_item_gap_hash->{$item}[$index_hrs_gap][$index_count]=$diff;
			}elsif($hrs_array[0][$index_count] ==0 or $hrs_array[$index_hrs][$index_count]==0){
				$diff=100;
			}else{
				$diff=($diff/$hrs_array[0][$index_count])*100;
			}
			$diff=sprintf("%.2f",$diff);
			#print $hrs_array[$index_hrs][$index_count],"\t",$hrs_array[0][$index_count],"\t",($hrs_array[$index_hrs][$index_count]-$hrs_array[0][$index_count]),"\t";
			#print $item,"\t",$diff,$/;
			#push @{$test_item_gap_hash->{$item}[$index_hrs_gap]},$diff;
			$test_item_gap_hash->{$item}[$index_hrs_gap][$index_count]=$diff;
		}
		$index_hrs_gap++;
	}
}
#print Dumper( $test_item_gap_hash);
#print Dumper($rel_hrs_array);
#print "ddd",$/;
my $saveName=$rel_file;
$saveName=~ s/(\..+$)/\.csv/;
print $saveName,$/;

open(OUT, ">$saveName") || die "**Error** can't open file [$saveName]: $!$/";

shift @$rel_hrs_array; #take out Ohr
print OUT "\t";
foreach my $item (@$test_item_array){
	#print $item,$/;
	#print "@{$test_item_gap_hash->{$item}}",$/;
	my @hrs_array=@{$test_item_gap_hash->{$item}[0]};
	#print @hrs_array,$/;
	for my $index_count (1 .. @hrs_array){
		print OUT $index_count,"\t";
	}
	print OUT $/;
	last;
}
foreach my $item (@$test_item_array){
	#print $item,"\t";
	#print $/;
	my $index_hrs=0;
	
	foreach my $hrs (@$rel_hrs_array){
		#print $item . '_' . $hrs,"\t";
		my @item_array;
		push @item_array,$item . '_' . $hrs;
		my $index=0;
		my @hrs_array=@{$test_item_gap_hash->{$item}[$index_hrs++]};
		my $flag=0;
		foreach my $eachdata (@hrs_array){
			if(abs($eachdata)>=$criterion){
				#print $eachdata;
				#print "\t";
				push @item_array,$eachdata;
				$flag=1;
			}else{
				#print $eachdata;
				#print "\t";
				if($print_all_data){
					push @item_array,$eachdata; 
				}else{
					push @item_array,"";
				}
			}
		}
		if($flag or $print_all_data){
			foreach my $eachdata (@item_array){
				print OUT $eachdata,"\t";
			}
			print OUT $/;
		}
	}

	
}
close(OUT); 