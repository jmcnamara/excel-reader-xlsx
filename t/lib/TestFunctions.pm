package TestFunctions;

###############################################################################
#
# TestFunctions - Helper functions for Excel::Reader::XLSX test cases.
#
# reverse('�'), January 2012, John McNamara, jmcnamara@cpan.org
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
    my $href;

    # Check if the JSON.pm module is avilable to parse the test data.
    eval { require JSON };

    if ( !$@ ) {

        # We have JSON.pm.
        my $json = JSON::XS->new();
        $href = $json->decode( $json_text );
    }
    else {

        # If JSON.pm isn't available we do a poor man's translation.
        $json_text =~ s/ : / => /g;
        $href = eval $json_text;
    }

    return $href;
}


1;


__END__

