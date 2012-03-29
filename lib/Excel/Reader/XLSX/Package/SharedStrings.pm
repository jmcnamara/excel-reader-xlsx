package Excel::Reader::XLSX::Package::SharedStrings;

###############################################################################
#
# SharedStrings - A class for reading the Excel XLSX sharedStrings.xml file.
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
use XML::LibXML::Reader qw(:types);
use Excel::Reader::XLSX::Package::XMLreader;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
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
    my $self  = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_count}        = 0;
    $self->{_unique_count} = 0;
    $self->{_strings}      = [];

    bless $self, $class;

    return $self;
}


##############################################################################
#
# _read_all_nodes()
#
# Override callback function. TODO rename.
#
sub _read_all_nodes {

    my $self = shift;
    my $reader = $self->{_reader};

    # Read the "shared string table" <sst> element for the count attributes.
    if ( $reader->nextElement( 'sst' ) ) {
        $self->{_count}        = $reader->getAttribute( 'count' );
        $self->{_unique_count} = $reader->getAttribute( 'uniqueCount' );
    }

    # Read the "string item" <si> elements.
    while ( $reader->nextElement( 'si' )  ) {

        my $string_node = $reader->copyCurrentNode( 1 );

        my $text = $string_node->textContent();

        push @{ $self->{_strings} }, $text;


       #    push @{ $self->{_strings} },
       #      [ $RICH_STRING, $string, $rich_string ];
    }
}


##############################################################################
#
# _read_rich_string()
#
# Read a rich string from an <si> element. A rich string is a string with
# multiple formats. The rich string is stored as a series of text "runs"
# denoted by <r> child elements. This function returns the raw string
# without formatting and the xml string with formatting.
#
sub _read_rich_string {

    my $self        = shift;
    my $node        = shift;
    my $string      = '';
    my $rich_string = '';

    # Get the nodes for the text runs <r>.
    for my $run_node ( $node->childNodes() ) {

        next unless $run_node->nodeName eq 'r';
        $rich_string .= $run_node->toString();

        # Get the nodes for the text <t>.
        for my $text_node ( $run_node->childNodes() ) {
            next unless $text_node->nodeName eq 't';
            $string .= $text_node->textContent();
        }
    }

    return ( $string, $rich_string );
}

###############################################################################
#
# _get_string()
#
# Get the shared string at the indexed value.
#
sub _get_string {

    my $self  = shift;
    my $index = shift;

    # Return an empty string is the index is out of bounds.
    return '' if $index < 0;
    return '' if $index >= $self->{_unique_count};

    my $string = $self->{_strings}->[$index];

    # For rich strings return the unformatted part of the string.
    if ( ref $string && $string->[0] == 1 ) {
        $string = $string->[1];
    }

    return $string;
}


1;


__END__

=pod

=head1 NAME

SharedStrings - A class for reading the Excel XLSX sharedStrings.xml file.

=head1 SYNOPSIS

See the documentation for L<Excel::Reader::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Reader::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

Copyright MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Reader::XLSX>.

=cut
