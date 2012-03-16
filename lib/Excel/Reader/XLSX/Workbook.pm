package Excel::Reader::XLSX::Workbook;

###############################################################################
#
# Workbook - A class for reading the Excel XLSX workbook.xml file.
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
use Excel::Reader::XLSX::Worksheet;
use Excel::Reader::XLSX::Package::Relationships;

our @ISA     = qw(Excel::Reader::XLSX::Package::XMLreader);
our $VERSION = '0.00';


###############################################################################
#
# Public and private API methods.
#
###############################################################################


###############################################################################
#
# new()
#
# Constructor.
#
sub new {

    my $class          = shift;
    my $package_dir    = shift;
    my $shared_strings = shift;
    my %files          = @_;

    my $self = Excel::Reader::XLSX::Package::XMLreader->new();

    $self->{_package_dir}          = $package_dir;
    $self->{_shared_strings}       = $shared_strings;
    $self->{_files}                = \%files;
    $self->{_worksheets}           = undef;
    $self->{_worksheet_attributes} = [];
    $self->{_worksheet_indexes}    = {};

    # Set the root dir for the workbook and worksheets. Usually 'xl/'.
    $self->{_workbook_root} = $self->{_files}->{_workbook};
    $self->{_workbook_root} =~ s/workbook.xml$//;

    bless $self, $class;

    $self->_set_relationships();

    return $self;
}


###############################################################################
#
# _set_relationships()
#
# Set up the relationship links between package files and the internal ids.
#
sub _set_relationships {

    my $self     = shift;
    my $filename = shift;

    my $rels_file = Excel::Reader::XLSX::Package::Relationship->new();

    $rels_file->_parse_file(
        $self->{_package_dir} . $self->{_files}->{_workbook_rels} );

    my %rels = $rels_file->_get_relationships();
    $self->{_rels} = \%rels;
}


##############################################################################
#
# _read_node()
#
# Callback function to read the nodes of the Workbook.xml file.
#
sub _read_node {

    my $self = shift;
    my $node = shift;

    # Only process the start elements.
    return unless $node->nodeType() == XML_READER_TYPE_ELEMENT;

    if ( $node->name eq 'sheet' ) {

        my $name     = $node->getAttribute( 'name' );
        my $sheet_id = $node->getAttribute( 'sheetId' );
        my $rel_id   = $node->getAttribute( 'r:id' );

        my $filename = $self->{_rels}->{$rel_id}->{_target};

        push @{ $self->{_worksheet_attributes} },
          {
            _name     => $name,
            _sheet_id => $sheet_id,
            _index    => $sheet_id - 1,
            _rel_id   => $rel_id,
            _filename => $filename,
          };
    }
}


###############################################################################
#
# worksheets()
#
# Return an array of Worksheet objects.
#
sub worksheets {

    my $self = shift;

    if ( !defined $self->{_worksheets} ) {
        $self->_read_worksheets();
    }

    return @{ $self->{_worksheets} };
}


###############################################################################
#
# worksheet_from_index()
#
# Return a Worksheet object based on its index.
#
sub worksheet_from_index {

    my $self  = shift;
    my $index = shift;

    return unless defined $index;

    if ( !defined $self->{_worksheets} ) {
        $self->_read_worksheets();
    }

    return $self->{_worksheets}->[$index];
}


###############################################################################
#
# Internal methods.
#
###############################################################################


###############################################################################
#
# _read_worksheets()
#
# Parse the workbook and set up the Worksheet objects.
#
sub _read_worksheets {

    my $self = shift;

    return if defined $self->{_worksheets};

    for my $sheet_attribute ( @{ $self->{_worksheet_attributes} } ) {

        my $worksheet = Excel::Reader::XLSX::Worksheet->new(
            $self->{_shared_strings},
            $sheet_attribute->{_name},
            $sheet_attribute->{_index},
        );

        $worksheet->_read_file(
                $self->{_package_dir}
              . $self->{_workbook_root}
              . $sheet_attribute->{_filename}

        );

        push @{ $self->{_worksheets} }, $worksheet;

        # Store the index so it can be looked up by name.
        $self->{_worksheets_index} = $sheet_attribute->{_name};
    }
}



1;


__END__

=pod

=head1 NAME

Workbook - A class for reading the Excel XLSX workbook.xml file.

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
