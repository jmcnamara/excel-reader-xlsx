###############################################################################
#
# Tests for Excel::Reader::XLSX::Package::SharedStrings methods.
#
# reverse('(c)'), January 2012, John McNamara, jmcnamara@cpan.org
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

$expected = 9;

$got = $reader->{_count};

_is_deep_diff( $got, $expected, $caption );


###############################################################################
#
# Test the _unique_count property.
#
$caption = " \tSharedStrings: _unique_count";

$expected = 7;

$got = $reader->{_unique_count};

_is_deep_diff( $got, $expected, $caption );



###############################################################################
#
# Test the _strings property.
#
$caption = " \tSharedStrings: _strings";

$expected = [
    'Foo', 'Bar', '  Foo', 'Foo ',
    'This is italic and this is bold',
    'This is italic and this is bold',
    '   This is italic and this is bold  '
];

$got = [];

for my $i (0 .. $reader->{_unique_count} -1) {
    $got->[$i] = $reader->_get_string( $i );
}


_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="9" uniqueCount="7"><si><t>Foo</t></si><si><t>Bar</t></si><si><t xml:space="preserve">  Foo</t></si><si><t xml:space="preserve">Foo </t></si><si><r><t xml:space="preserve">This is </t></r><r><rPr><i/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>italic</t></r><r><rPr><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve"> and this is </t></r><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>bold</t></r></si><si><r><rPr><i/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>This is italic</t></r><r><rPr><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve"> </t></r><r><rPr><i/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>and this is bold</t></r></si><si><r><rPr><i/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">   This is </t></r><r><rPr><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">italic </t></r><r><rPr><i/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">and this is </t></r><r><rPr><b/><i/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve">bold  </t></r></si></sst>
