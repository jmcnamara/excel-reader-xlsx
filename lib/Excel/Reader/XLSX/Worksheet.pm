package Excel::Reader::XLSX::Worksheet;

###############################################################################
#
# Worksheet - A class for reading the Excel XLSX sheet.xml file.
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
    };

    bless $self, $class;


    #$self->_set_relationships();

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
# Read all the nodes of a worksheet.xml file using an XML::LibXML::Reader
# instance.
#
sub _read_all_nodes {

    my $self = shift;

    # TODO remove.

    while ( $self->{_reader}->read() ) {
        $self->_read_node( $self->{_reader} );
    }
}

###############################################################################
#
# next_row()
#
# TODO
#
sub next_row {

    my $self = shift;

    my $has_node = $self->{_reader}->nextElement( 'row' );

    if ( $has_node ) {
        $self->{_row_is_empty} = $self->{_reader}->isEmptyElement();
        $self->{_end_of_row}   = 0;
    }


    return $has_node;
}


###############################################################################
#
# next_cell()
#
# TODO
#
sub next_cell {

    my $self = shift;
    my $cell;
    my $node;
    my $cell_start = 0;

    return if $self->{_row_is_empty};
    return if $self->{_end_of_row};



    while ( !$cell_start ) {

        return if !$self->{_reader}->read();
        $node = $self->{_reader};

        if (   $node->name() eq 'c'
            && $node->nodeType() == XML_READER_TYPE_ELEMENT )
        {
            $cell_start = 1;
            last;
        }

        if ( $node->name eq 'row' ) {
            $self->{_end_of_row} = 1;
            return;
        }


    }


    my $range = $node->getAttribute( 'r' );
    return unless $range;

    ( $cell->{_row}, $cell->{_col} ) = _range_to_rowcol( $range );


    my $type = $node->getAttribute( 't' ) || '';

    $cell->{_type} = $type;


    my $cell_node = $node->copyCurrentNode( $FULL_DEPTH );


    # Read the cell <c> child nodes.
    for my $child_node ( $cell_node->childNodes() ) {

        if ( $child_node->nodeName() eq 'v' ) {
            $cell->{_value}     = $child_node->textContent();
            $cell->{_has_value} = 1;
        }
        elsif ( $child_node->nodeName() eq 'f' ) {
            $cell->{_formula}     = $child_node->textContent();
            $cell->{_has_formula} = 1;
        }
    }


    use Data::Dumper::Perltidy;
    print Dumper $cell;

    return $cell;
}




###############################################################################
#
# name()
#
# TODO
#
sub name {

    my $self = shift;

    return $self->{_name};
}





###############################################################################
#
# _range_to_rowcol($range)
#
# TODO.
#
sub _range_to_rowcol {

    my $range = shift;

    $range =~ /([A-Z]{1,3})(\d+)/;

    my $col = $1;
    my $row = $2;

    # Convert base26 column string to number.
    my @chars = split //, $col;
    my $exponent = 0;
    $col = 0;

    while ( @chars ) {
        my $char = pop @chars;    # LS char first
        $col += ( ord( $char ) - ord( 'A' ) + 1 ) * ( 26**$exponent );
        $exponent++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    return $row, $col;
}





1;


__END__

=pod

=head1 NAME

Worksheet - A class for reading the Excel XLSX sheet.xml file.

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
