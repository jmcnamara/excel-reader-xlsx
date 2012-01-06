package Excel::Reader::XLSX::Workbook;

###############################################################################
#
# Workbook - A class for reading the Excel XLSX workbook.xml file.
#
# Used in conjunction with Excel::Reader::XLSX
#
# Copyright 2012, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

# perltidy with the following options: -mbl=2 -pt=0 -nola

use 5.008002;
use strict;
use warnings;
use Exporter;
use Carp;
use XML::LibXML::Reader;
use Excel::Reader::XLSX::Worksheet;
use Excel::Reader::XLSX::Package::Relationships;

our @ISA     = qw(Exporter);
our $VERSION = '0.00';

our $FULL_DEPTH  = 1;
our $RICH_STRING = 1;


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;

    my $package_dir = shift;
    my %files       = @_;

    my $self = {
        _reader          => undef,
        _package_dir     => $package_dir,
        _files           => \%files,
        _worksheets      => undef,
        _worksheet_attributes => [],
    };

    bless $self, $class;


    $self->_set_relationships();

    return $self;
}


##############################################################################
#
# _read_file()
#
# Create an XML::LibXML::Reader instance from a file.
#
sub _read_file {

    my $self     = shift;
    my $filename = shift;

    my $xml_reader = XML::LibXML::Reader->new( location => $filename );

    $self->{_reader} = $xml_reader;

    return $xml_reader;
}


##############################################################################
#
# _read_string()
#
# Create an XML::LibXML::Reader instance from a string. Used mainly for 
# testing.
#
sub _read_string {

    my $self   = shift;
    my $string = shift;

    my $xml_reader = XML::LibXML::Reader->new( string => $string );

    $self->{_reader} = $xml_reader;

    return $xml_reader;
}


##############################################################################
#
# _read_filehandle()
#
# Create an XML::LibXML::Reader instance from a filehandle. Used mainly for
# testing.
#
sub _read_filehandle {

    my $self       = shift;
    my $filehandle = shift;

    my $xml_reader = XML::LibXML::Reader->new( IO => $filehandle );

    $self->{_reader} = $xml_reader;

    return $xml_reader;
}


##############################################################################
#
# _read_all_nodes()
#
# Read all the nodes of a workbook.xml file using an XML::LibXML::Reader
# instance.
#
sub _read_all_nodes {

    my $self = shift;

    while ( $self->{_reader}->read() ) {
        $self->_read_node( $self->{_reader} );
    }
}


###############################################################################
#
# _set_relationships()
#
# TODO
#
sub _set_relationships {

    my $self     = shift;
    my $filename = shift;

    my $rels_file = Excel::Reader::XLSX::Package::Relationship->new();

    $rels_file->_read_file(
        $self->{_package_dir} . 'xl/_rels/workbook.xml.rels' );

    $rels_file->_read_all_nodes();

    my %rels = $rels_file->_get_relationships();
    $self->{_rels} = \%rels;
}





##############################################################################
#
# _read_node()
#
# Callback function to read the nodes of the Workbook.xml file.
#
sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only process the start elements.
    return unless $node->nodeType() == XML_READER_TYPE_ELEMENT;


    if ( $node->name eq 'sheet' ) {

        my $name     = $node->getAttribute( 'name' );
        my $sheet_id = $node->getAttribute( 'sheetId' );
        my $rel_id   = $node->getAttribute( 'r:id' );

        my $filename = $self->{_rels}->{$rel_id}->{_target};


          push @{ $self->{_worksheet_attributes} },
          {
            _name     => $name,
            _sheet_id => $sheet_id,
            _rel_id   => $rel_id,
            _filename => $filename,
          };
    }
}

###############################################################################
#
# worksheets()
#
# TODO
#
sub worksheets {

    my $self = shift;

    if ( defined $self->{_worksheets} ) {
        return @{ $self->{_worksheets} };
    }

    for my $attribute ( @{ $self->{_worksheet_attributes} } ) {
        print ">> $attribute->{_filename}\n";

        my $worksheet = Excel::Reader::XLSX::Worksheet->new();

        $worksheet->{_name} = $attribute->{_name};


        $worksheet->_read_file(
            $self->{_package_dir} . 'xl/'. $attribute->{_filename} );

        push @{ $self->{_worksheets} }, $worksheet;
    }

    return @{ $self->{_worksheets} };
}





1;


__END__

=pod

=head1 NAME

Workbook - A class for reading the Excel XLSX workbook.xml file.

=head1 SYNOPSIS

See the documentation for L<Excel::Reader::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Reader::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Reader::XLSX>.

=cut
