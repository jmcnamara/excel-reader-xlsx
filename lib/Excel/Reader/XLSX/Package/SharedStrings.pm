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
# _read_node()
#
# Callback function to read the nodes of the <sst> shared string table data.
#
sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only process the start elements.
    return unless $node->nodeType() == XML_READER_TYPE_ELEMENT;

    # Read the "shared string table" <sst> element for the count attributes.
    if ( $node->name() eq 'sst' ) {
        $self->{_count}        = $node->getAttribute( 'count' );
        $self->{_unique_count} = $node->getAttribute( 'uniqueCount' );
    }

    # Read the "string item" <si> elements.
    if ( $node->name() eq 'si' ) {

        my $string_node = $node->copyCurrentNode( $FULL_DEPTH );

        # Read the <si> child nodes.
        for my $text_node ( $string_node->childNodes() ) {

            if ( $text_node->nodeName() eq 't' ) {

                # Read a plain text node.
                push @{ $self->{_strings} }, $text_node->textContent();
                last;
            }
            elsif ( $text_node->nodeName() eq 'r' ) {

                # Read a plain rich text node.
                my ( $string, $rich_string ) =
                  _read_rich_string( $self, $string_node );

                push @{ $self->{_strings} },
                  [ $RICH_STRING, $string, $rich_string ];
                last;
            }
        }
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

© MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Reader::XLSX>.

=cut
