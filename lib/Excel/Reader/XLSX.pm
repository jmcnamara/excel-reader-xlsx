package Excel::Reader::XLSX;

###############################################################################
#
# WriteExcelXML.
#
# Excel::Reader::XLSX - Read data from an Excel 2007+/XLSX format file.
#
# Copyright 2012, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use 5.008002;
use strict;
use warnings;
use Exporter;
use Archive::Zip;
use OLE::Storage_Lite;
use File::Temp qw(tempdir);
use Excel::Reader::XLSX::Workbook;
use Excel::Reader::XLSX::Package::ContentTypes;
use Excel::Reader::XLSX::Package::SharedStrings;


# Modify Archive::Zip error handling reporting so we can catch errors.
Archive::Zip::setErrorHandler( sub { die shift } );

our @ISA     = qw(Exporter);
our $VERSION = '0.00';

# Error codes for some common errors.
our $ERROR_none                      = 0;
our $ERROR_file_not_found            = 1;
our $ERROR_file_is_xls               = 2;
our $ERROR_file_is_encrypted         = 3;
our $ERROR_file_is_unknown_ole       = 4;
our $ERROR_file_zip_error            = 5;
our $ERROR_file_missing_subfile      = 6;
our $ERROR_file_has_no_content_types = 7;
our $ERROR_file_has_no_workbook      = 8;

our @error_strings = (
    '',                                   # 0
    'File not found',                     # 1
    'File is xls not xlsx',               # 2
    'File is encrypted xlsx',             # 3
    'File is unknown OLE doc type',       # 4
    'File has zip error',                 # 5
    'File missing subfile',               # 6
    'File has no [Content_Types].xml',    # 7
    'File has no workbook.xml',           # 8
);


###############################################################################
#
# new()
#
sub new {

    my $class = shift;

    my $self = {
        _reader           => undef,
        _files            => {},
        _tempdir          => undef,
        _error_status     => 0,
        _error_extra_text => '',
    };

    bless $self, $class;

    return $self;
}


###############################################################################
#
# read_file()
#
# Unzip the XLSX file and read the [Content_Types].xml file to get the
# structure of the contained XML files.
#
# Return a valid Workbook object if sucessful. If not return undef and set
# the error status.
#
sub read_file {

    my $self     = shift;
    my $filename = shift;

    # Check that the file exists.
    if ( !-e $filename ) {
        $self->{_error_status}     = $ERROR_file_not_found;
        $self->{_error_extra_text} = $filename;
        return;
    }


    # Check for xls or encrypted OLE files.
    my $ole_file = $self->_check_if_ole_file( $filename );
    if ( $ole_file ) {
        $self->{_error_status}     = $ole_file;
        $self->{_error_extra_text} = $filename;
        return;
    }

    # Create a, locally scoped, temp dir to unzip the XLSX file into.
    my $tempdir = File::Temp->newdir( DIR => $self->{_tempdir} );

    # Archive::Zip requires a Unix directory separator to the end.
    $tempdir .= '/' if $tempdir !~ m{/$};


    # Create an Archive::Zip object to unzip the XLSX file.
    my $zipfile = Archive::Zip->new();

    # Read the XLSX zip file and catch any errors.
    eval { $zipfile->read( $filename ) };

    # Store the zip error and return.
    if ( $@ ) {
        my $error_text = $@;
        chomp $error_text;
        $self->{_error_status}     = $ERROR_file_zip_error;
        $self->{_error_extra_text} = $error_text;
        return;
    }

    # Extract the XML files from the XLSX zip.
    $zipfile->extractTree( '', $tempdir );

    # The [Content_Types] is required as the root of the other files.
    my $content_types_file = $tempdir . '[Content_Types].xml';

    if ( !-e $content_types_file ) {
        $self->{_error_status} = $ERROR_file_has_no_content_types;
        return;
    }

    # Create a reader object to read the [Content_Types].
    my $content_types = Excel::Reader::XLSX::Package::ContentTypes->new();
    $content_types->_read_file( $content_types_file );
    $content_types->_read_all_nodes();

    # Read the filenames from the [Content_Types].
    my %files = $content_types->_get_files();

    # Check that the listed files actually exist.
    my $files_exist = $self->_check_files_exist( $tempdir, %files );

    if ( !$files_exist ) {
        $self->{_error_status} = $ERROR_file_missing_subfile;
        return;
    }

    # Verify that the workbook.xml file is listed.
    if ( !$files{_workbook} ) {
        $self->{_error_status} = $ERROR_file_has_no_workbook;
        return;
    }

    # Create a reader object to read the sharedStrings.xml file.
    my $shared_strings = Excel::Reader::XLSX::Package::SharedStrings->new();

    # Read the sharedStrings if present. Only files with strings have one.
    if ( $files{_shared_strings} ) {

        $shared_strings->_read_file( $tempdir . $files{_shared_strings} );
        $shared_strings->_read_all_nodes();
    }

    # Create a reader object for the workbook.xml file.
    my $workbook = Excel::Reader::XLSX::Workbook->new(
        $tempdir,
        $shared_strings,
        %files

    );

    # Read data from the workbook.xml file.
    $workbook->_read_file( $tempdir . $files{_workbook} );
    $workbook->_read_all_nodes();

    # Store information in the reader object.
    $self->{_files}          = \%files;
    $self->{_shared_strings} = $shared_strings;
    $self->{_package_dir}    = $tempdir;
    $self->{_zipfile}        = $zipfile;

    return $workbook;
}


###############################################################################
#
# _check_files_exist()
#
# Verify that the subfiles read from the Content_Types actually exist;
#
sub _check_files_exist {

    my $self    = shift;
    my $tempdir = shift;
    my %files   = @_;
    my @filenames;

    # Get the filenames for the files hash.
    for my $key ( keys %files ) {
        my $filename = $files{$key};

        # Worksheets are stored in an aref.
        if ( ref $filename ) {
            push @filenames, @$filename;
        }
        else {
            push @filenames, $filename;
        }
    }

    # Verify that the files exist.
    for my $filename ( @filenames ) {
        if ( !-e $tempdir . $filename ) {
            $self->{_error_extra_text} = $filename;
            return;
        }
    }

    return 1;
}


###############################################################################
#
# _check_if_ole_file()
#
# Check if the file in an OLE compound doc. This can happen in a few cases.
# This first is when the file is xls and not xlsx. The second is when the
# file is an encrypted xlsx file. We also handle the case of unknown OLE
# file types.
#
# Porting note. As a lightweight test you can check for OLE files by looking
# for the magic number 0xD0CF11E0 (docfile0) at the start of the file.
#
sub _check_if_ole_file {

    my $self     = shift;
    my $filename = shift;
    my $ole      = OLE::Storage_Lite->new( $filename );
    my $pps      = $ole->getPpsTree();

    # If getPpsTree() failed then this isn't an OLE file.
    return if !$pps;

    # Loop throught the PPS children below the root.
    for my $child_pps ( @{ $pps->{Child} } ) {

        my $pps_name = OLE::Storage_Lite::Ucs2Asc( $child_pps->{Name} );

        # Match an Excel xls file.
        if ( $pps_name eq 'Workbook' || $pps_name eq 'Book' ) {
            return $ERROR_file_is_xls;
        }

        # Match an encrypted Excel xlsx file.
        if ( $pps_name eq 'EncryptedPackage' ) {
            return $ERROR_file_is_encrypted;
        }
    }

    return $ERROR_file_is_unknown_ole;
}


###############################################################################
#
# error().
#
# Return an error string for a failed read.
#
sub error {

    my $self        = shift;
    my $error_index = $self->{_error_status};
    my $error       = $error_strings[$error_index];

    if ( $self->{_error_extra_text} ) {
        $error .= ': ' . $self->{_error_extra_text};
    }

    return $error;
}


###############################################################################
#
# error_code().
#
# Return an error code for a failed read.
#
sub error_code {

    my $self = shift;

    return $self->{_error_status};
}


1;


__END__



=head1 NAME

Excel::Reader::XLSX - Read data from an Excel 2007+/XLSX format file.

=head1 SYNOPSIS

TODO.

=head1 DESCRIPTION

TODO.


=head1 TODO

=over

=item * Reading from filehandles.

=back


=head1 DISCLAIMER OF WARRANTY

Because this software is licensed free of charge, there is no warranty for the software, to the extent permitted by applicable law. Except when otherwise stated in writing the copyright holders and/or other parties provide the software "as is" without warranty of any kind, either expressed or implied, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose. The entire risk as to the quality and performance of the software is with you. Should the software prove defective, you assume the cost of all necessary servicing, repair, or correction.

In no event unless required by applicable law or agreed to in writing will any copyright holder, or any other party who may modify and/or redistribute the software as permitted by the above licence, be liable to you for damages, including any general, special, incidental, or consequential damages arising out of the use or inability to use the software (including but not limited to loss of data or data being rendered inaccurate or losses sustained by you or third parties or a failure of the software to operate with any other software), even if such holder or other party has been advised of the possibility of such damages.




=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.




=head1 AUTHOR

John McNamara jmcnamara@cpan.org




=head1 COPYRIGHT

Copyright MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.


