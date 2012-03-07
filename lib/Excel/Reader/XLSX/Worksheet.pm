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
use Carp;
use Excel::Reader::XLSX::Package::XMLreader;
use Excel::Reader::XLSX::Row;


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

    $self->{_shared_strings}      = shift;
    $self->{_previous_row_number} = -1;

    bless $self, $class;

    return $self;
}


###############################################################################
#
# next_row()
#
# Read the next available row in the worksheet.
#
sub next_row {

    my $self = shift;
    my $row  = undef;

    my $has_row = $self->{_reader}->nextElement( 'row' );

    if ( $has_row ) {

        my $node = $self->{_reader};
        my $row_number = $node->getAttribute( 'r' );

        if ( defined $row_number ) {

            # Convert to zero indexed value.
            $row_number--;
        }
        else {

            # If no 'r' attribute assume it is one more than the previous.
            $row_number = $self->{_previous_row_number} + 1;
        }

        $row = Excel::Reader::XLSX::Row->new(

            $self->{_reader},
            $self->{_shared_strings},
            $row_number,
            $self->{_previous_row_number},
        );

        $self->{_previous_row_number} = $row_number;
    }

    return $row;
}


###############################################################################
#
# name()
#
# Return the worksheet name.
#
sub name {

    my $self = shift;

    return $self->{_name};
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

Copyright MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Reader::XLSX>.

=cut
