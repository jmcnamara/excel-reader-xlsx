###############################################################################
#
# Tests for Excel::Writer::XLSX.
#
# reverse('©'), February 2012, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_is_deep_diff _read_json);
use strict;
use warnings;
use Excel::Reader::XLSX;

use Test::More tests => 1;

###############################################################################
#
# Test setup.
#
my $json_filename = 't/regression/json_files/read_rows04.json';
my $json          = _read_json( $json_filename );
my $caption       = $json->{caption};
my $expected      = $json->{expected};
my $xlsx_file     = 't/regression/xlsx_files/' . $json->{xlsx_file};
my $got;


###############################################################################
#
# Test reading data from an Excel file.
#
use Excel::Reader::XLSX;

my $reader   = Excel::Reader::XLSX->new();
my $workbook = $reader->read_file( $xlsx_file );

for my $worksheet ( $workbook->worksheets() ) {

    my $sheetname = $worksheet->name();
    $got->{$sheetname} = [];

    while ( my $row = $worksheet->next_row() ) {

        my @values = $row->values();
        push @{ $got->{$sheetname} }, [@values];

        # Reread row to get cached data.
        @values = $row->values();
        push @{ $got->{$sheetname} }, [@values];
    }
}


# Test the results.
_is_deep_diff( $got, $expected, $caption );
