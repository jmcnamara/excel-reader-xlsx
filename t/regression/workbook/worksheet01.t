###############################################################################
#
# Tests for Excel::Writer::XLSX.
#
# reverse('(c)'), February 2012, John McNamara, jmcnamara@cpan.org
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
my $json_filename = 't/regression/json_files/workbook/worksheet01.json';
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

# Check for valid indexes.
for my $index ( -1, -2, 0, 1 ) {

    my $worksheet = $workbook->worksheet( $index );

    my $sheetname = $worksheet->name();
    push @$got, $sheetname;
}

# Check for invalid indexes.
for my $index ( -5, 4 ) {

    my $worksheet = $workbook->worksheet( $index );

    next unless $worksheet;

    my $sheetname = $worksheet->name();
    push @$got, $sheetname;
}


# Test the results.
_is_deep_diff( $got, $expected, $caption );
