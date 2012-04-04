package Excel::Reader::XLSX::Cell;

###############################################################################
#
# Cell - A class for reading the Excel XLSX cells.
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
use Excel::Reader::XLSX::Package::XMLreader;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class = shift;
    my $self  = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_shared_strings} = shift;
    $self->{_value}          = '';

    bless $self, $class;

    return $self;
}


###############################################################################
#
# _init()
#
# Initialise a Cell object.
#
sub _init {

    my $self = shift;

    $self->{_value}            = '';
    $self->{_converted_string} = 0;
    $self->{_has_formula}      = 0;
}


###############################################################################
#
# value()
#
# Return the cell value.
#
sub value {

    my $self = shift;

    # If the cell type is a shared string convert the value index to a string.
    if ( $self->{_type} eq 's' && !$self->{_converted_string} ) {
        $self->{_value} =
          $self->{_shared_strings}->_get_string( $self->{_value} );

        # State variable so that multiple calls to value() don't need lookups.
        $self->{_converted_string} = 1;
    }


    return $self->{_value};
}


###############################################################################
#
# row()
#
# Return the cell row number, zero-indexed.
#
sub row {

    my $self = shift;

    return $self->{_row};
}


###############################################################################
#
# col()
#
# Return the cell column number, zero indexed.
#
sub col {

    my $self = shift;

    return $self->{_col};
}


1;


__END__

=pod

=head1 NAME

Cell - A class for reading the Excel XLSX cells.

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
