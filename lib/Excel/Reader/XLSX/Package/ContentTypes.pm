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
use XML::LibXML::Reader qw(:types);
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

    $self->{_files} = {};

    bless $self, $class;

    return $self;
}


##############################################################################
#
# _read_node()
#
# Callback function to read the <Types> attributes of the ContentTypes file.
# We currently only read files/types that we are interested in.
#
sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only read the Override nodes.
    return unless $node->name eq 'Override';


    my $part_name    = $node->getAttribute( 'PartName' );
    my $content_type = $node->getAttribute( 'ContentType' );


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
# Get a hash of of the files read from the ContentTypes file.
#
sub _get_files {

    my $self = shift;

    return %{ $self->{_files} };

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

Copyright MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.

=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.

=head1 DISCLAIMER OF WARRANTY

See the documentation for L<Excel::Reader::XLSX>.

=cut
