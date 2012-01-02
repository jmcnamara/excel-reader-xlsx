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

use strict;
use Excel::Reader::XLSX::Workbook;

our @ISA     = qw(Excel::Reader::XLSX::Workbook Exporter);
our $VERSION = '0.00';


###############################################################################
#
# new()
#
sub new {

    my $class = shift;
    my $self  = Excel::Reader::XLSX::Workbook->new( @_ );

    # Check for file creation failures before re-blessing
    bless $self, $class if defined $self;

    return $self;
}


1;


__END__



=head1 NAME

Excel::Reader::XLSX - Read data from an Excel 2007+/XLSX format file.

=head1 SYNOPSIS

TODO.

=head1 DESCRIPTION

TODO.



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


