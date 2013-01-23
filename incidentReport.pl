#!/usr/bin/perl                                                                                                                                                                  
$| = 1;                                                                                                                                                                          
use strict;
use Getopt::Std;
use Data::Dumper;

my @custArray = (); # fill with all the group codes. 

sub set_group_array() {

  my @stuff_ref = @_ ;

  my $cust_array ;
  my $current_customer = '';

 for my $foo (@stuff_ref) {
   $cust_array =  @{$foo}[1];
#   print "$cust_array\n";
   my @cust_split = split('\.', $cust_array);
#  print "$cust_split[1] \n"; 
   
   my $next = "$cust_split[1]";  
#  print "Starting cust: $current_customer\n";

   if ("$current_customer" ne "$next" ) {
     $current_customer = $next;
     print "current_customer: $current_customer \n";
     print "next: $next\n" ; 
     push(@custArray, $current_customer);
   }
  
 }
  print "@custArray\n";
}




sub generateCustomerOutageArray () {
my @customerArray = @_;

   my $startTime            = '';
   my @returnArray          = ();
   my @customerArrayElement = ();

   #
   # Gather data from start of the outage to the end of the data array
   my @bufferArray  = ();
   my $outageStart  = '';
   foreach my $customerElement ( @customerArray ) {
      @customerArrayElement = @{$customerElement};
      ## print "@customerArrayElement";
      ## print "\n";
      if ( $customerArrayElement[5] < 66 && ! $outageStart ) {
         $outageStart = $customerArrayElement[0];
         push(@bufferArray, [ @customerArrayElement ] );
         ## print "@customerArrayElement";
         ## print "\n";
      } elsif ( $outageStart ) {
         push(@bufferArray, [ @customerArrayElement ] );
         ## print "@customerArrayElement";
         ## print "\n";
      }
   }

   #
   # Gather data from end of the data array to the alleged end of the outage 
   @customerArray = @bufferArray;
   @bufferArray   = ();
   my @last       = ();
   my $outageEnd  = '';
   foreach my $customerElement ( reverse(@customerArray) ) {
      @customerArrayElement = @{$customerElement};
      if ( @last ) {
         if ( ( ( $customerArrayElement[5] < 60 && $last[5] > 60 ) || ($last[5] < 60) ) && ! $outageEnd ) {
            $outageEnd = $customerArrayElement[0];
            push(@bufferArray, [ @last ] );
            push(@bufferArray, [ @customerArrayElement ] );
         } elsif ( $outageEnd ) {
            push(@bufferArray, [ @customerArrayElement ] );
         } else {
            @last = @customerArrayElement;
         }
      } else {
         @last = @customerArrayElement;
         if ( $last[5] < 60 ) {
            $outageEnd = $last[0];
            push(@bufferArray, [ @last ] );
         }
      }
   }
   @returnArray = reverse(@bufferArray);
   return(@returnArray);
}

my @descriptionFull = ();

sub descriptionNote() {
  print "Please type in a description of the incident: \n" ;
  my $description = <>;
  chomp ($description);
  my $descriptionHeader = ' ==DESCRIPTION== ';
  @descriptionFull = ( $descriptionHeader, $description );
}


my @impactFull = ();

sub impactNote() {
  print "Enter the Impact it had: \n";
  my $impact = <>;
  chomp($impact);
  my $impactHeader = ' ==IMPACT== ';
  @impactFull = ( $impactHeader, $impact );
}

my @commentFull = ();

sub commentNote() {
  print "Enter any Comments or Observations you would like to note: \n";
  my $comment = <>;
  chomp($comment);
  my $commentHeader = ' ==COMMENTS/OBSERVATIONS== ';
  @commentFull = ( $commentHeader, $comment );
}

my @custArray = ();
my $custName = '';

sub writeXLSFile () {
my @sorted = @_;
use Spreadsheet::WriteExcel;

   my @header            = ();
   my @customerHeader    = ('Event Time', 'Customer Name', 'Start Time', 'End Time', 'Unsuccessful Tests', 'Total Pages Tested', '% Unsuccessful', 'Calculated Unavailability');

   # Create a new Excel WorkBook
   my $workbook = Spreadsheet::WriteExcel->new('/tmp/incidentReport.xls');

   # Add a New WorkSheet
   my $reportWorkSheet = $workbook->add_worksheet('Incident Report');
   $reportWorkSheet->set_header('&CGenerated at &T');
   $reportWorkSheet->set_landscape();
   $reportWorkSheet->set_paper(1);
   $reportWorkSheet->set_margins_LR('.5');
   $reportWorkSheet->center_horizontally();
   $reportWorkSheet->set_first_sheet();
   $reportWorkSheet->activate();

   # Raw Data
   my $rawRowNumber = 0;
   my $workSheet = $workbook->add_worksheet('RawData');
   $workSheet->set_column('A:A', 30);
   $workSheet->set_column('B:B', 50);
   $workSheet->set_column('C:L', 20);

   my @rowArrayElement = ();

   my @currentCustomerArray  = ();
   my $currentCustomerRow    = 0;
   my $currentCustomer       = '';

   my $podId                 = '';
   my @workSheetNames        = ();
   my $customerWorkSheet     = '';
   my $customerWorkSheetName = '';
   foreach my $rowElement ( @sorted ){
      @rowArrayElement = @{$rowElement};
      ## print "@rowArrayElement";
      ## print "\n";

      #
      # We Need to Break out the Customer Name because of the length Limitation of the Worksheet Name.
      #
      if ( $currentCustomer ne $rowArrayElement[1] ) {
         if ( $currentCustomer ) {
            # New Customer Grouping Else First Customer
            @currentCustomerArray = &generateCustomerOutageArray(@currentCustomerArray);
            $customerWorkSheet->write_col('A4', \@currentCustomerArray);
            $rawRowNumber++;
         }

         $currentCustomer  = $rowArrayElement[1];
         my (@arrayBuffer) = split(/\./, $currentCustomer);
         if ( $arrayBuffer[1] && $arrayBuffer[2] && $arrayBuffer[3] ) {
            @currentCustomerArray = ();
            $currentCustomerRow   = 0;
 
            $customerWorkSheetName = sprintf("%s-%s-%s", $arrayBuffer[1], $arrayBuffer[2], $arrayBuffer[3]);
            $customerWorkSheet     = $workbook->add_worksheet($customerWorkSheetName);

            $customerWorkSheet->set_column('A:A', 30);
            $customerWorkSheet->set_column('B:B', 50);
            $customerWorkSheet->set_column('C:L', 20);
            $customerWorkSheet->write_row($currentCustomerRow, 0, \@customerHeader);
            $currentCustomerRow++;
            $customerWorkSheet->write_formula($currentCustomerRow,0,'=(COUNTA(A4:A500)*5)-5');
            $customerWorkSheet->write_formula($currentCustomerRow,1,'=B4');
            $customerWorkSheet->write_formula($currentCustomerRow,2,'=A4');
            $customerWorkSheet->write_formula($currentCustomerRow,3,'=INDEX(A4:A5000,COUNTA(A4:A5000),1,1)');
            $customerWorkSheet->write_formula($currentCustomerRow,4,'=SUM(E4:E5000)');
            $customerWorkSheet->write_formula($currentCustomerRow,5,'=SUM(G4:G5000)');
            $customerWorkSheet->write_formula($currentCustomerRow,6,'=E2/F2');
            $customerWorkSheet->write_formula($currentCustomerRow,7,'=G2*A2');
            $currentCustomerRow++;
            $customerWorkSheet->write_row($currentCustomerRow, 0, \@header);
            $currentCustomerRow++;

            if ( ! $podId ) {
               $podId = uc($arrayBuffer[0]);
            }
            push(@currentCustomerArray, [ @rowArrayElement ] );
            push(@workSheetNames, $customerWorkSheetName);
         } else {
            # No Data Meand we Found the Header Line
            $currentCustomer = '';
            @header          = @rowArrayElement;
         }
      } else {
         push(@currentCustomerArray, [ @rowArrayElement ] );
      }

      # Everything gets Written to rawData
      $workSheet->write_row($rawRowNumber, 0, \@rowArrayElement);
      $rawRowNumber++;
   }
   if ( $currentCustomer ) {
      # Last Customer Grouping
      @currentCustomerArray = &generateCustomerOutageArray(@currentCustomerArray);
      $customerWorkSheet->write_col('A4', \@currentCustomerArray);
   }

   #
   # Generate Incident Report Summary
   # $reportWorkSheet->set_column('A:A', 60);
   my @reportHeader = ('Realm', 'Start Time', 'End Time', 'Calculated Unavailability (Minutes)');
   $reportWorkSheet->set_column('A:A', 50);
   $reportWorkSheet->set_column('B:C', 20);
   $reportWorkSheet->set_column('D:D', 32);
   $reportWorkSheet->set_row(0, 30);
   $reportWorkSheet->set_row(2, 30);
   my $big_font = $workbook->add_format(size => 20);
   my $med_font = $workbook->add_format(size => 15);
   my $bold     = $workbook->add_format();
   $bold->set_bold();
   my $title = sprintf("Demandware Incident Report: %s", $podId);
   $reportWorkSheet->write_string('A1',$title,$big_font);
   $reportWorkSheet->write_string('A3',"Realms Affected",$med_font);
   $reportWorkSheet->write_string('C3',"Type:",$med_font);
   $reportWorkSheet->write_string('D3',"UnPlanned",$med_font);
   $reportWorkSheet->data_validation('D3',
                                     {
                                      input_title   => 'Planned or UnPlanned Event',
                                      validate => 'list',
                                      value    => ['Planned', 'UnPlanned'],
                                     });
   $reportWorkSheet->write_row(4, 0, \@reportHeader, $bold);

   my $reportRow = 5;
   my $formula = '';
   foreach my $worksheet (@workSheetNames) {
      $formula = sprintf "=%s!B2", $worksheet;
      $reportWorkSheet->write_formula($reportRow,0,$formula);
      $formula = sprintf "=%s!C2", $worksheet;
      $reportWorkSheet->write_formula($reportRow,1,$formula);
      $formula = sprintf "=%s!D2", $worksheet;
      $reportWorkSheet->write_formula($reportRow,2,$formula);
      $formula = sprintf "=%s!H2", $worksheet;
      $reportWorkSheet->write_formula($reportRow,3,$formula);
      $reportRow++;
   }
   
   descriptionNote();
   impactNote();
   commentNote();
   
   my @commentArray = ();
   
   if ( $descriptionFull[1] ne '' ) {
	   push(@commentArray, @descriptionFull);
   }
   if ( $impactFull[1] ne '' ) {
	   push(@commentArray, @impactFull);
   }
   if ( $commentFull[1] ne '' ) {
	   push(@commentArray, @commentFull);
   }
   
   $reportRow++; # to add a blank row between
   foreach my $note (@commentArray) {
	   $reportRow++;
	   $reportWorkSheet->write_string($reportRow, 0, $note);
   }
   
   $workbook->close();
}

sub processInputFile() {
use Text::CSV;
my ($infile) = (@_);
my @returnArray = ();

   if (!open(INFILE,$infile)) {
      print "\n\tError: Unable open datafile\n\n";
   } else {
      my $csv = Text::CSV->new({ sep_char => ',' });
      my $row = 0;
      my @rowArray = (); 
      while (my $line = <INFILE>) {
         if ($csv->parse($line)) {
            push(@{$rowArray[$row]},$csv->fields());
            $row++;
         } else {
            printf "Line Could Not Be Parsed:\n%s\n", $line;
         }
      }
      close(INFILE);
 
      @returnArray = sort { $a->[1] cmp $b->[1] || $a->[0] cmp $b->[0] } @rowArray;

   }
   return(@returnArray);
}

MAIN:
{
my $infile = "";
my @returnArray = ();
my %opts = ();
my $exitCode = 1;
my $Usage = "
        Usage:  incidentReport.pl -f <FileName>

";

   if (!getopts("f:", \%opts)) {
      printf "\nInvalid Option Sprcified\n\n";
      $exitCode = 1;
   } else {
      if ( $opts{f} ) {
         if ( ! -e $opts{f}  ) {
            printf "\nError: Unable open specified file $opts{f}\n\n";
            $exitCode = 1;
         } else {
            $infile = $opts{f};
         }
      }
   }

   if ( $infile ) {
      @returnArray = &processInputFile($infile);
      &writeXLSFile(@returnArray);
#   print Dumper(@returnArray);
#     &set_group_array($infile);
      &set_group_array(@returnArray);
   } else {
      print "$Usage";
   }
   exit($exitCode);
}

