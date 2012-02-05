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
  _read_json
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


###############################################################################
#
# _read_json()
#
# Read test data from a JSON file.
#
sub _read_json {

    my $filename = shift;

    # Read in the JSON test data
    local $/;
    open my $fh, '<:encoding(UTF-8)', $filename
      or die "Couldn't open $filename\n";

    my $json_text = <$fh>;


    use JSON;
    my $data      = decode_json( $json_text );

    return $data;
}



1;


__END__

