#!/usr/bin/perl -w

use warnings;
use strict; 
use Unicode::Normalize;		
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Spreadsheet::Read;
use Spreadsheet::WriteExcel;

#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 1: INPUT EXCEL FILENAME AND ESTABLISH ACCESS
#-------------------------------------------------------------------------------------------------------------------------------------#
print "\n..................................................\n        Y MAZE ALTERNATION DATA PROCESSOR        \n..................................................\n
	\nBefore we begin, I need to ask you five questions.\n\nWhich file should I open?\n";

my $stdin = <STDIN>; 		# defines my $stdin variable as the user input and converts it to lowercase
	chomp ($stdin);		# I guess extra characters are added when you do stdin?  This removes them.

sleep 1;

if (-e $stdin){
}	
else{
print "\n\nI CAN'T OPEN THE FILE YOU SPECIFIED\nThe file name that you typed does not exist. 
	\nPlease try running the program again using\na corrected filename with .xlsx or .xls \nextension.\n\n\n";
	sleep 5;
	exit;
	}
	
	
my $fileinput = ReadData($stdin);
my @array = ();
my @temp = ();
my @acells=();
my @bcells=();
my @ccells=();
my @dcells=();
my @ecells=();
my @fcells=();
my @cells=();
my @alts= ();
my @length=();
my @possiblealts= ();
my @percentalts= ();
my @cellarray= ();
my @namecellarray=();
my @names=();




#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 2: INITIALIZE ALL POSSIBLE CELL NUMBERS, A1-A50 THROUGH F1-F50
#-------------------------------------------------------------------------------------------------------------------------------------#

my $a="a";
my $b="b";
my $c="c";
my $d="d";
my $e="e";
my $f="f";


my $generalnumber = 50;
for (my $i = 0; $i < $generalnumber; $i++){	
	my $x = ($i+1);
	my $acell=uc join('',$a,$x);
	chomp ($acell);
	push (@acells, $acell);
	
	my $bcell=uc join('',$b,$x);
	chomp ($bcell);
	push (@bcells, $bcell);
	
	my $ccell=uc join('',$c,$x);
	chomp ($ccell);
	push (@ccells, $ccell);
	
	my $dcell=uc join('',$d,$x);
	chomp ($dcell);
	push (@dcells, $dcell);
	
	my $ecell=uc join('',$e,$x);
	chomp ($ecell);
	push (@ecells, $ecell);
	
	my $fcell=uc join('',$f,$x);
	chomp ($fcell);
	push (@fcells, $fcell);
}


#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 3: PROMPT USER TO INPUT THE RANGE OF ROWS. THIS IS NEEDED FOR ALL LOOPS THAT THAT USE A SPECIFIC RANGE OF CELL NUMBERS.
#-------------------------------------------------------------------------------------------------------------------------------------#

print "\nGreat! I can access that file.\n"; sleep 1;
print "These next questions are about your excel data file.\n"; sleep 1;
print "\nWhich row number contains:\n...the first line of data?    ";	
my $firstrownumber = <STDIN>; 				#Determine which row number to begin generating cells.
	chomp ($firstrownumber);


if ($firstrownumber =~ /[0-9]/) {
	}
else {
	print "\nYou did not enter a number.\n"; sleep 1;
	print "Please rerun program and enter a number next time.\n";
	sleep 5;
	exit;
}		


print "...the last line of data?     ";
my $almostlastrownumber = <STDIN>; 				#Determine which row number to stop generating cells for.
	chomp ($almostlastrownumber);	
	
my $lastrownumber=($almostlastrownumber+1);

if ($lastrownumber =~ /[0-9]/) {
	}
else {
	print "\nYou did not enter a number.\n"; sleep 1;
	print "Please rerun program and enter a number next time.\n";
	sleep 5;
	exit;
}		

my $totalnumberofrows=($lastrownumber-$firstrownumber); #This is needed later on, to specify how many times (ie rows) $i needs to 
							# run a loop for.



#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 4: ASK WHICH COLUMN THE MOUSE NAMES ARE IN.  BY KNOWING THE COLUMNS (step 4) AND ROWS (step 3), CREATE ARRAY OF THESE CELL No.
#-------------------------------------------------------------------------------------------------------------------------------------#

print "\nWhich column contains:\n...the mouse IDs?             ";
my $namecolumn = <STDIN>; 				#Determine which column contains the name data.
	chomp ($namecolumn);

if ($namecolumn =~ /[0-9]/) {
	print "\nYou did not enter a letter.\n"; sleep 1;
	print "Please rerun program and enter a letter next time.\n";
	sleep 5;
	exit;}
else {
}	

for (my $i = $firstrownumber; $i < $lastrownumber; $i++){	#Starting at the specified first row number, keep running through the data
								#so long as $i is less than the specified last row number.
	my $namecells=uc join('',$namecolumn,$i);	#Joins the loop number ($i) with the column letter to create a list of cell
							#numbers which contain the name data.
	chomp ($namecells);
	push (@namecellarray, $namecells);		#'Push' adds the generated cells to the array 

}



#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 5: ASK WHICH COLUMN THE ARM ENTRY DATA IS IN.  BY KNOWING THE COLUMNS (step 5) AND ROWS (step 3), CREATE ARRAY OF THESE CELL No.
#-------------------------------------------------------------------------------------------------------------------------------------#

print "...the arm entry data?        ";
my $column = <STDIN>; 					#Determine which column contains the arm entry data
	chomp ($column);

if ($column =~ /[0-9]/) {
	print "\nYou did not enter a letter.\n"; sleep 1;
	print "Please rerun program and enter a letter next time.\n";
	sleep 5;
	exit;}
else {
}	

print "\n";

for (my $i = $firstrownumber; $i < $lastrownumber; $i++){	#Starting at the specified first row number, keep running through the data	
								#so long as $i is less than the specified last row number.	
	my $columnandcell=uc join('',$column,$i);	#Joins the loop number ($i) with the column letter to create a list of cell
							#numbers which contain the name data.
	push (@cellarray, $columnandcell);		#'Push' adds the generated cells to the array 
}
	


#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 6: USE LOOP TO ADD ALL DATA FROM CERTAIN CELLS TO ARRAY
#-------------------------------------------------------------------------------------------------------------------------------------#
for (my $i = 0; $i < $totalnumberofrows; $i++){		#starting at place 0, so long as i is less than the
							#user inputted number, then go to the next one

	if ($fileinput->[1]{$cellarray[$i]} =~ /^[+-]?\d+$/ ) { #this goes through the array created in 5 (lists all cells that contain
								#arm alternations)  So long as the input from these cells is a number,
	push (@array, $fileinput->[1]{$cellarray[$i]});		#it will add that cell/rows arm visit data to an array (@array)
	push (@names, $fileinput->[1]{$namecellarray[$i]});	#and it will add the corresponding 'name' cell to an array (@names)
	} 
}




#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 7.1: FOR EACH VALUE OF ARRAY, SPLIT STRING INTO ITS OWN ARRAY
#-------------------------------------------------------------------------------------------------------------------------------------#
foreach my $array (@array){ 			#for every string in our excel import array

if ($array ne 0){				#if the (current) string is non-zero
	my @temp=split(//,$array);		#split the string and make it an array 

	my $length = length ($array);		#we need to define the number of elements of each arm entry data set.
						#given that each mouse will likely enter different arms a different number of times,
						#we need to store the total arm visits in an array. this is needed a) for raw data and
						#b) for the alternation calculations
	push (@length, $length);		
my $successfulalternation = 0;
my $possiblealternations = $length-2;



#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 7.2: FOR EACH SPLIT STRING, USE LOOP TO CALCULATE HOW MANY ALTERNATIONS OCCURED AND PUSH ALTERNATION VALUES INTO NEW ARRAY
#-------------------------------------------------------------------------------------------------------------------------------------#
for (my $i = 0; $i < $length-2; $i++){		#starting at place 0, so long as the final spot is before (total length-2), keep going
	if ($temp[$i] ne $temp[$i+1] && $temp[$i] ne $temp[$i+2] ){	# if (where you are at in the array) doesn't 
									# equal the next value AND doesn't equal the following
	$successfulalternation +=1; 					# value, then tally that as a successful alternation
	}								# END if statement (individual strings' characters in array)
}									# END for loop (looping through individual strings' characters)
	push (@alts, $successfulalternation);				
	push (@possiblealts, $possiblealternations);
	}								# END if statement (if current chosen value in array is non-zero)
}									#END foreach (for each value in array)


my $actuallength = scalar(@alts);


# RECAP: For each value in my array (each value is a cell imported from excel), if the value is non-blank, split the value into its
# individual characters and make an array of the characters.  So, this takes data like '12321' and breaks it into '1','2','3','2','1'
# and puts each character in an array. This is necessary because the alternation calculator runs a for loop through this character
# array.



print "\nAlternation data:"; sleep 1;
print " Calculated.";
print "\nPossible alternations:"; sleep 1;
print " Calculated.";




#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 8: USING VALUES FROM @ALTS ARRAY AND @POSSIBLEALTS, CALCULATE PERCENT ALTERNATION AND PUSH THESE VALUES INTO NEW ARRAY
#-------------------------------------------------------------------------------------------------------------------------------------#


for (my $i = 0; $i < $actuallength; $i++){
    push (@percentalts, ($alts[$i]/$possiblealts[$i]*100));
}

print "\nPercentage alternations:"; sleep 1;
print " Calculated."; sleep 1;



#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 9: CREATE NEW EXCEL FILE, PER THE USER INPUTTED NAME
#-------------------------------------------------------------------------------------------------------------------------------------#
print "\n\n\nWhat should I save the new data file as?\nPlease include the extension (.xls, not .xlsx)\n";

 my $savefile = <STDIN>; 		# defines my $stdin variable as the user input and converts it to lowercase
	chomp ($savefile);		# I guess extra characters are added when you do stdin?  This removes them.


print "\nThe new data will be saved as $savefile, \nin the same folder as this program.\n\n";
 
 
    my $workbook = Spreadsheet::WriteExcel->new($savefile);   	#creates a new excel file, named per the user input
    my $worksheet = $workbook->add_worksheet();               	#adds a workshet to your new excel file
    my $format = $workbook->add_format();			#lets you format the sheets?
    my $bold = $workbook->add_format();
    my $format01 = $workbook->add_format();
	$bold->set_bold();
	$bold->set_align('center');
	$bold->set_border(1);
	$bold->set_font('arial');
	$bold->set_size(9);
	
	$format->set_align('center');
	$format->set_border(1);
	$format->set_font('arial');
	$format->set_size(9);

	$format01->set_align('center');
	$format01->set_border(1);
	$format01->set_font('arial');
	$format01->set_size(9);
	$format01->set_num_format('00.00"%"');
  	
	$worksheet->set_column( 'A:F', 19 );			#sets the column width to 18
	$worksheet->write('A1', 'Mouse Numbers', $bold);	#titles header and makes it bold
	$worksheet->write('B1', 'Total Arm Visits', $bold);	#titles header and makes it bold
	$worksheet->write('C1', 'Possible Alternations', $bold);#titles header and makes it bold
	$worksheet->write('D1', 'Actual Alternations', $bold);	#titles header and makes it bold
	$worksheet->write('E1', 'Percent Alternations', $bold);	#titles header and makes it bold


#-------------------------------------------------------------------------------------------------------------------------------------#
# STEP 10: EXPORT DATA FROM ARRAYS (possible alts, total alts, percent alt) INTO THE EXCEL FILE
#-------------------------------------------------------------------------------------------------------------------------------------#

for (my $i = 0; $i < $actuallength; $i++){	#starting at place 0, so long as i is less than the
						#user inputted number, then go to the next one
	my $x = ($i+1);
								#we generated all cells A1-A50 through F1-F50 before. 
	$worksheet->write($acells[$x], $names[$i], $format);		#this puts all items from the corresponding array (@names, @length, etc)
	$worksheet->write($bcells[$x], $length[$i], $format); 		#into the new excel file, starting at 'xcell' position 1 
	$worksheet->write($ccells[$x], $possiblealts[$i], $format); 	#note: since arrays start at place zero instead of 1, I had to define
	$worksheet->write($dcells[$x], $alts[$i], $format); 		#x as i+1.
	$worksheet->write($ecells[$x], $percentalts[$i], $format01); 

	} 


sleep 2;

print "\n\nCONGRATS!  Your data is done processing.  You can open your file now.\n\n\n";

#-------------------------------------------------------------------------------------------------------------------------------------#