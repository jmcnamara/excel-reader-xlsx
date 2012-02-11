package Excel::Reader::XLSX::Package::ContentTypes;

###############################################################################
#
# ContentTypes - A class for reading the Excel XLSX ContentTypes.xml file.
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

our @ISA     = qw(Exporter);
our $VERSION = '0.00';


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;

    my $self = {
        _reader       => undef,
        _files        => {},
    };

    bless $self, $class;

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
# Read all the nodes of a ContentTypes.xml file using an XML::LibXML::Reader
# instance.
#
sub _read_all_nodes {

    my $self = shift;

    while ( $self->{_reader}->read() ) {
        $self->_read_node( $self->{_reader} );
    }
}


##############################################################################
#
# _read_node()
#
# Callback function to read the <Types> attributes of the ContentTypes file.
#
sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only read the Override nodes.
    return unless $node->name eq 'Override';


    my $part_name    = $node->getAttribute('PartName');
    my $content_type = $node->getAttribute('ContentType');


    # Strip leading directory separator from filename.
    $part_name =~ s{^/}{};

    if ( $part_name =~ /app\.xml$/ ) {
        $self->{_files}->{_app} = $part_name;
        return;
    }

    if ( $part_name =~ /core\.xml$/ ) {
        $self->{_files}->{_core} = $part_name;
        return;
    }

    if ( $part_name =~ /sharedStrings\.xml$/ ) {
        $self->{_files}->{_shared_strings} = $part_name;
        return;
    }

    if ( $part_name =~ /styles\.xml$/ ) {
        $self->{_files}->{_styles} = $part_name;
        return;
    }

    if ( $part_name =~ /workbook\.xml$/ ) {

        # The workbook.xml.rels file isn't included in the ContentTypes but
        # it is usually in the _rels dir at the same level at the workbook.xml.
        my $workbook_rels = $part_name;
        $workbook_rels =~ s{(workbook.xml)}{_rels/$1.rels};

        $self->{_files}->{_workbook}      = $part_name;
        $self->{_files}->{_workbook_rels} = $workbook_rels;
        return;
    }

    if ( $part_name =~ /sheet\d+\.xml$/ ) {
        push @{ $self->{_files}->{_worksheets} }, $part_name;
        return;
    }
}


###############################################################################
#
# _get_files()
#
# TODO
#
sub _get_files {

    my $self = shift;

    return %{$self->{_files}};

}




1;


__END__

=pod

=head1 NAME

ContentTypes - A class for reading the Excel XLSX ContentTypes.xml file.

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
