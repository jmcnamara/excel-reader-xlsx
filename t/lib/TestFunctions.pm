package TestFunctions;

###############################################################################
#
# TestFunctions - Helper functions for Excel::Reader::XLSX test cases.
#
# reverse('©'), January 2012, John McNamara, jmcnamara@cpan.org
#

use 5.008002;
use Exporter;
use strict;
use warnings;
use Test::More;


our @ISA         = qw(Exporter);
our @EXPORT      = ();
our %EXPORT_TAGS = ();
our @EXPORT_OK   = qw(
  _is_deep_diff
);

our $VERSION = '0.00';


###############################################################################
#
# Use Test::Differences::eq_or_diff() where available or else fall back to
# using Test::More::is_deeply().
#
sub _is_deep_diff {
    my ( $got, $expected, $caption, ) = @_;

    eval {
        require Test::Differences;
        Test::Differences->import();
    };

    if ( !$@ ) {
        eq_or_diff( $got, $expected, $caption, { context => 1 } );
    }
    else {
        is_deeply( $got, $expected, $caption );
    }

}


1;


__END__

