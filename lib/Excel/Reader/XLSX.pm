package Excel::Reader::XLSX;

###############################################################################
#
# WriteExcelXML.
#
# Excel::Reader::XLSX - Efficient data reader for the Excel XLSX file format.
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
use File::Temp;
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
    'File is missing subfile',            # 6
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
# Return a valid Workbook object if successful. If not return undef and set
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
    $content_types->_parse_file( $content_types_file );

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

        $shared_strings->_parse_file( $tempdir . $files{_shared_strings} );
    }

    # Create a reader object for the workbook.xml file.
    my $workbook = Excel::Reader::XLSX::Workbook->new(
        $tempdir,
        $shared_strings,
        %files

    );

    # Read data from the workbook.xml file.
    $workbook->_parse_file( $tempdir . $files{_workbook} );

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

    # Loop through the PPS children below the root.
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

Excel::Reader::XLSX - Efficient data reader for the Excel XLSX file format.

=head1 SYNOPSIS

The following is a simple Excel XLSX file reader using C<Excel::Reader::XLSX>:

    use strict;
    use warnings;
    use Excel::Reader::XLSX;

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );

    if ( !defined $workbook ) {
        die $reader->error(), "\n";
    }

    for my $worksheet ( $workbook->worksheets() ) {

        my $sheetname = $worksheet->name();

        print "Sheet = $sheetname\n";

        while ( my $row = $worksheet->next_row() ) {

            while ( my $cell = $row->next_cell() ) {

                my $row   = $cell->row();
                my $col   = $cell->col();
                my $value = $cell->value();

                print "  Cell ($row, $col) = $value\n";
            }
        }
    }

    __END__



=head1 DESCRIPTION

C<Excel::Reader::XLSX> is a fast and lightweight parser for Excel XLSX files. XLSX is the Office Open XML, OOXML, format used by Excel 2007 and later.

B<Note: This software is designated as alpha quality until this notice is removed.> The API shouldn't change but functionality is currently limited.

=head1 Reader

The C<Excel::Reader::XLSX> constructor returns a Reader object that is used to read an Excel XLSX file:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    die $reader->error() if !defined $workbook;

    for my $worksheet ( $workbook->worksheets() ) {
        while ( my $row = $worksheet->next_row() ) {
            while ( my $cell = $row->next_cell() ) {
                my $value = $cell->value();
                ...
            }
        }
    }

The C<Excel::Reader::XLSX> object is used to return sub-objects that represent the functional parts of an Excel spreadsheet, L</Workbook>, L</Worksheet>, L</Row> and L</Cell>:

     Reader
       +- Workbook
          +- Worksheet
             +- Row
                +- Cell

The C<Reader> object has the following methods:

    read_file()
    error()
    error_code()

=head2 read_file()

The C<read_file> Reader method is used to read an Excel XLSX file and return a C<Workbook> object:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    ...

It is recommended that the success of the C<read_file()> method is always checked using one of the error checking methods below.

=head2 error()

The C<error()> Reader method returns an error string if C<read_file()> fails:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );

    if ( !defined $workbook ) {
        die $reader->error(), "\n";
    }
    ...

The C<error()> strings and associated C<error_code()> numbers are:

    error()                              error_code()
    =======                              ============
    ''                                   0
    'File not found'                     1
    'File is xls not xlsx'               2
    'File is encrypted xlsx'             3
    'File is unknown OLE doc type'       4
    'File has zip error'                 5
    'File is missing subfile'            6
    'File has no [Content_Types].xml'    7
    'File has no workbook.xml'           8


=head2 error_code()

The C<error_code()> Reader method returns an error code if C<read_file()> fails:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );

    if ( !defined $workbook ) {
        die "Got error code ", $parser->error_code, "\n";
    }

This method is useful if you wish to use you own error strings or error handling methods.


=head1 Workbook

=head2 Workbook Methods

An C<Excel::Reader::XLSX> C<Workbook> object is returned by the Reader C<read_file()> method:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    ...

The C<Workbook> object has the following methods:

    worksheets()
    worksheet()

=head2 worksheets()

The Workbook C<worksheets()> method returns an array of
C<Worksheet> objects. This method is generally used to iterate through
all the worksheets in an Excel workbook and read the data:

    for my $worksheet ( $workbook->worksheets() ) {
      ...
    }


=head2 worksheet()

The Workbook C<worksheet()> method returns a single C<Worksheet>
object using the sheetname or the zero based index.

    my $worksheet = $workbook->worksheet( 'Sheet1' );

    # Or via the index.

    my $worksheet = $workbook->worksheet( 0 );


=head1 Worksheet

=head2 Worksheet Methods

The C<Worksheet> object is returned from a L</Workbook> object and is used to access row data.

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    die $reader->error() if !defined $workbook;

    for my $worksheet ( $workbook->worksheets() ) {
        ...
    }

The C<Worksheet> object has the following methods:

     next_row()
     name()
     index()

=head2 next_row()

The C<next_row()> method returns a L</Row> object representing the next
row in the worksheet.

        my $row = $worksheet->next_row();

It returns C<undef> if there are no more rows containing data or formatting in the worksheet. This allows you to iterate over all the rows in a worksheet as follows:

        while ( my $row = $worksheet->next_row() ) { ... }

Note, for efficiency the C<next_row()> method returns the next row in the file. This may not be the next sequential row. An option to read sequential rows, wheter they contain data or not will be added in a later release.

=head2 name()

The C<name()> method returns the name of the Worksheet object.

    my $sheetname = $worksheet->name();

=head2 index()

The C<index()> method returns the zero-based index of the Worksheet
object.

    my $sheet_index = $worksheet->index();


=head1 Row

=head2 Row Methods

The C<Row> object is returned from a L</Worksheet> object and is use to access cells in the worksheet.

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    die $reader->error() if !defined $workbook;

    for my $worksheet ( $workbook->worksheets() ) {
        while ( my $row = $worksheet->next_row() ) {
            ...
        }
    }

The C<Row> object has the following methods:

    values()
    next_cell()
    row_number()


=head2 values()

The C<values())> method returns an array of values for a row from the first column up to the last column containing data. Cells with no data value return an empty string C<''>.

    my @values = $row->values();

For example if we extracted data for the first row of the following spreadsheet we would get the values shown below:

     -----------------------------------------------------------
    |   |     A     |     B     |     C     |     D     | ...
     -----------------------------------------------------------
    | 1 |           | Foo       |           | Bar       | ...
    | 2 |           |           |           |           | ...
    | 3 |           |           |           |           | ...

    # Code:
    ...
    my $row = $worksheet->next_row();
    my @values = $row->values();
    ...

    # @values contains ( '', 'Foo', '', 'Bar' )


=head2 next_cell()

The C<next_cell> method returns the next, non-blank cell in the current row.

    my $cell = $row->next_cell();

It is usually used with a while loop. For example if we extracted data for the first row of the following spreadsheet we would get the values shown below:

     -----------------------------------------------------------
    |   |     A     |     B     |     C     |     D     | ...
     -----------------------------------------------------------
    | 1 |           | Foo       |           | Bar       | ...
    | 2 |           |           |           |           | ...
    | 3 |           |           |           |           | ...

    # Code:
    ...
    while ( my $cell = $row->next_cell() ) {
        my $value = $cell->value();
        print $value, "\n";
    }
    ...

    # Output:
    Foo
    Bar

Note, for efficiency the C<next_cell()> method returns the next cell in the row. This may not be the next sequential cell. An option to read sequential cells, wheter they contain data or not will be added in a later release.


=head2 row_number()

The C<row_number()> method returns the zero-indexed row number for the current row:

    my $row = $worksheet->next_row();
    print $row->row_number(), "\n";


=head1 Cell

=head2 Cell Methods

The C<Cell> object is used to extract data from Excel cells:

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );
    die $reader->error() if !defined $workbook;

    for my $worksheet ( $workbook->worksheets() ) {
        while ( my $row = $worksheet->next_row() ) {
            while ( my $cell = $row->next_cell() ) {
                my $value = $cell->value();
               ...
            }
        }
    }

The C<Cell> object has the following methods:

    value()
    row()
    col()

For example if we extracted the data for the cells in the first row of the following spreadsheet we would get the values shown below:

     -----------------------------------------------------------
    |   |     A     |     B     |     C     |     D     | ...
     -----------------------------------------------------------
    | 1 |           | Foo       |           | Bar       | ...
    | 2 |           |           |           |           | ...
    | 3 |           |           |           |           | ...

    # Code:
    ...
    while ( my $row = $worksheet->next_row() ) {
        while ( my $cell = $row->next_cell() ) {
            my $row   = $cell->row();
            my $col   = $cell->col();
            my $value = $cell->value();

            print "Cell ($row, $col) = $value\n";
        }
    }
    ...

    # Output:
    Cell (0, 1) = Foo
    Cell (0, 2) = Bar


=head2 value()

The Cell C<value()> method returns the unformatted value from the cell.

    my $value = $cell->value();

The "value" of the cell can be a string or  a number. In the case of a formula it returns the result of the formula and not the formal string. For dates it returns the numeric serial date.


=head2 row()

The Cell C<row()> method returns the zero-indexed row number of the cell.

    my $row = $cell->row();


=head2 col()

The Cell C<col()> method returns the zero-indexed column number of the cell.

    my $col = $cell->col();


=head1 EXAMPLE

Simple example of iterating through all worksheets in a workbook and printing out values from cells that contain data.

    use strict;
    use warnings;
    use Excel::Reader::XLSX;

    my $reader   = Excel::Reader::XLSX->new();
    my $workbook = $reader->read_file( 'Book1.xlsx' );

    if ( !defined $workbook ) {
        die $reader->error(), "\n";
    }

    for my $worksheet ( $workbook->worksheets() ) {

        my $sheetname = $worksheet->name();

        print "Sheet = $sheetname\n";

        while ( my $row = $worksheet->next_row() ) {

            while ( my $cell = $row->next_cell() ) {

                my $row   = $cell->row();
                my $col   = $cell->col();
                my $value = $cell->value();

                print "  Cell ($row, $col) = $value\n";
            }
        }
    }

=head1 RATIONALE

The rationale for this module is to have a fast memory efficient module for reading XLSX files. This is based on my experience of user requirements as the maintainer of Spreadsheet::ParseExcel.


=head1 SEE ALSO

Spreadsheet::XLSX, an XLSX reader using the old Spreadsheet::ParseExcel hash based interface: L<http://search.cpan.org/dist/Spreadsheet-XLSX/>.

SimpleXlsx, a "rudimentary extension to allow parsing of information stored in Microsoft Excel XLSX spreadsheets": L<http://search.cpan.org/dist/SimpleXlsx/>.

Excel::Writer::XLSX, an XLSX file writer based on the Spreadsheet::WriteExcel interface: L<http://search.cpan.org/dist/Excel-Writer-XLSX/>.


=head1 TODO

There are a lot of features still to be added. This module is very much a work in progress.

=over

=item * Reading from filehandles.

=item * Option to read sequential rows via C<next_row()>.

=item * Option to read dates instead of raw serial style numbers. This is actually harder than it would seem due to the XLSX format.

=item * Option to read formulas, urls, comments, images.

=item * Spreadsheet::ParseExcel style interface.

=item * Direct cell access.

=item * Cell format data.

=back




=head1 LICENSE

Either the Perl Artistic Licence L<http://dev.perl.org/licenses/artistic.html> or the GPL L<http://www.opensource.org/licenses/gpl-license.php>.




=head1 AUTHOR

John McNamara jmcnamara@cpan.org




=head1 COPYRIGHT

Copyright MMXII, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.




=head1 DISCLAIMER OF WARRANTY

Because this software is licensed free of charge, there is no warranty for the software, to the extent permitted by applicable law. Except when otherwise stated in writing the copyright holders and/or other parties provide the software "as is" without warranty of any kind, either expressed or implied, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose. The entire risk as to the quality and performance of the software is with you. Should the software prove defective, you assume the cost of all necessary servicing, repair, or correction.

In no event unless required by applicable law or agreed to in writing will any copyright holder, or any other party who may modify and/or redistribute the software as permitted by the above licence, be liable to you for damages, including any general, special, incidental, or consequential damages arising out of the use or inability to use the software (including but not limited to loss of data or data being rendered inaccurate or losses sustained by you or third parties or a failure of the software to operate with any other software), even if such holder or other party has been advised of the possibility of such damages.
