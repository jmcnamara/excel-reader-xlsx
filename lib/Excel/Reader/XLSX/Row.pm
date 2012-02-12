package Excel::Reader::XLSX::Row;

###############################################################################
#
# Row - A class for reading Excel XLSX rows.
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
use Carp;
use XML::LibXML::Reader;
use Excel::Reader::XLSX::Cell;
use Excel::Reader::XLSX::Package::XMLreader;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';

our $FULL_DEPTH = 1;


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $self  = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_reader}         = shift;
    $self->{_shared_strings} = shift;
    $self->{_row_number}     = shift;
    $self->{_row_is_empty}   = $self->{_reader}->isEmptyElement();
    $self->{_end_of_row}     = 0;

    bless $self, $class;

    return $self;
}


###############################################################################
#
# next_cell()
#
# Get the cell cell in the current row.
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

    # Create a cell object.
    $cell = Excel::Reader::XLSX::Cell->new( $self->{_shared_strings} );


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
        if ( $child_node->nodeName() eq 'is' ) {
            $cell->{_value}     = $child_node->textContent();
            $cell->{_has_value} = 1;
        }
        elsif ( $child_node->nodeName() eq 'f' ) {
            $cell->{_formula}     = $child_node->textContent();
            $cell->{_has_formula} = 1;
        }
    }


    return $cell;
}


###############################################################################
#
# number()
#
# Return the cell row number, zero-indexed.
#
sub number {

    my $self = shift;

    return $self->{_row_number};
}


#
# Internal methods.
#

###############################################################################
#
# _range_to_rowcol($range)
#
# Convert an Excel A1 style ref to a zero indexed row and column.
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

Row - A class for reading Excel XLSX rows.

=head1 SYNOPSIS

See the documentation for L<Excel::Reader::XLSX>.

=head1 DESCRIPTION

This module is used in conjunction with L<Excel::Reader::XLSX>.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

� MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Reader::XLSX>.

=cut
