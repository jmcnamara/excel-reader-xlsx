###############################################################################
#
# Tests for Excel::Reader::XLSX::Package::ContentTypes methods.
#
# reverse('(c)'), January 2012, John McNamara, jmcnamara@cpan.org
#

use lib 't/lib';
use TestFunctions qw(_is_deep_diff);
use strict;
use warnings;
use Excel::Reader::XLSX::Package::ContentTypes;

use Test::More tests => 1;

###############################################################################
#
# Tests setup.
#
my $expected;
my $caption;
my $got;
my $reader = Excel::Reader::XLSX::Package::ContentTypes->new();

$reader->_read_filehandle( *DATA );
$reader->_read_all_nodes();



###############################################################################
#
# Test the _files property.
#
$caption = " \tContentTypes: _strings";

$expected = {
    '_workbook'      => 'xl/workbook.xml',
    '_workbook_rels' => 'xl/_rels/workbook.xml.rels',
    '_app'           => 'docProps/app.xml',
    '_styles'        => 'xl/styles.xml',
    '_worksheets'    => [
        'xl/worksheets/sheet1.xml',
        'xl/worksheets/sheet2.xml',
        'xl/worksheets/sheet3.xml'

    ],
    '_core'           => 'docProps/core.xml',
    '_shared_strings' => 'xl/sharedStrings.xml'
};


$got = $reader->{_files};

_is_deep_diff( $got, $expected, $caption );

__DATA__
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>
