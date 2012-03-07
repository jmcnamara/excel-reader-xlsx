package Excel::Reader::XLSX::Package::Relationship;

###############################################################################
#
# Relationship - A class for reading the Excel XLSX Rels file.
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

    $self->{_rels} = {};

    bless $self, $class;

    return $self;
}


##############################################################################
#
# _read_node()
#
# Callback function to read the <Types> attributes of the Relationship file.
#
sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only read the Override nodes.
    return unless $node->name eq 'Relationship';

    my $id          = $node->getAttribute( 'Id' );
    my $type        = $node->getAttribute( 'Type' );
    my $target      = $node->getAttribute( 'Target' );
    my $target_mode = $node->getAttribute( 'TargetMode' );

    $self->{_rels}->{$id} = {
        _type        => $type,
        _target      => $target,
        _target_mode => $target_mode,
    };
}


###############################################################################
#
# _get_relationships()
#
# Return a hash to the relationships.
#
sub _get_relationships {

    my $self = shift;

    return %{ $self->{_rels} };
}


1;


__END__

=pod

=head1 NAME

Relationship - A class for reading the Excel XLSX Rels file.

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
