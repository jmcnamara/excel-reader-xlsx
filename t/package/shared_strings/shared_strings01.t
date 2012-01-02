###############################################################################
#
# Tests for Excel::Reader::XLSX::Package::SharedStrings methods.
#
# reverse('(c)'), September 2010, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_is_deep_diff);
use strict;
use warnings;
use Excel::Reader::XLSX::Package::SharedStrings;

use Test::More tests => 3;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $reader = Excel::Reader::XLSX::Package::SharedStrings->new();

$reader->_read_filehandle( *DATA );
$reader->_read_all_nodes();



###############################################################################
#
# Test the _count property.
#
$caption = " \tSharedStrings: _count";

$expected = 7;

$got = $reader->{_count};

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test the _unique_count property.
#
$caption = " \tSharedStrings: _unique_count";

$expected = 3;

$got = $reader->{_unique_count};

_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test the _strings property.
#
$caption = " \tSharedStrings: _strings";

$expected = [ 'neptune', 'mars', 'venus' ];

$got = $reader->{_strings};

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">
  <si>
    <t>neptune</t>
  </si>
  <si>
    <t>mars</t>
  </si>
  <si>
    <t>venus</t>
  </si>
</sst>
