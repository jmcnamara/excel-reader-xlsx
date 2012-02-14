#!/usr/bin/perl

use ExtUtils::MakeMaker;
use strict;
use warnings;

require 5.008002;

my %deps = (
    'Archive::Zip'        => 1.30,
    'XML::LibXML'         => 1.89,
    'OLE::Storage_Lite'   => 0.19,
);

my %resources = (
    homepage    => 'http://jmcnamara.github.com/excel-reader-xlsx/',
    repository  => 'http://github.com/jmcnamara/excel-reader-xlsx',
    bugtracker  => 'http://github.com/jmcnamara/excel-reader-xlsx/issues',
    MailingList => 'http://groups.google.com/group/spreadsheet-writeexcel',
);

my %params = (
    'AUTHOR'        => 'John McNamara (jmcnamara@cpan.org)',
    'ABSTRACT'      => 'Read data from an Excel 2007+/XLSX format file.',
    'NAME'          => 'Excel::Reader::XLSX',
    'VERSION_FROM'  => 'lib/Excel/Reader/XLSX.pm',
    'NEEDS_LINKING' => 0,
    'dist'          => { COMPRESS => 'gzip --best', SUFFIX => 'gz' },
    'PREREQ_PM'     => \%deps,
    'test'          => { TESTS => 't/*/*.t t/*/*/*.t' }

);

if ( $ExtUtils::MakeMaker::VERSION ge '6.46' ) {
    $params{META_MERGE} = { resources => \%resources };
}

if ( $ExtUtils::MakeMaker::VERSION ge '6.30_01' ) {
    $params{LICENSE} = 'perl';
}

WriteMakefile(%params);

# Add other targets to the makefile.
package MY;

sub postamble {
"
localdoc : \$(VERSION_FROM)
	pod2cpanhtml \$(VERSION_FROM) > ~/tmp/pod2html.html
"
}



