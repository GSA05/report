use strict;
use File::Copy;
#use Data::Dumper;
#use re "debug";

my $args = join " ", @ARGV;
my %mpi = (
    using => 0,
    begin => 1,
    end   => 1,
);
my %omp = (
    using => 0,
    begin => 1,
    end   => 1,
);
my @model = ( 16, 20, 32 );
if ( $args =~ /-mpi( (\d{1,3})(:(\d{1,3}))?|$| )/ ) {
    $mpi{'using'} = 1;
    $mpi{'begin'} = $2 || 1;
    $mpi{'end'}   = $4 || $2 || 1;
}
if ( $args =~ /-omp( (\d{1,3})(:(\d{1,3}))?|$| )/ ) {
    $omp{'using'} = 1;
    $omp{'begin'} = $2 || 1;
    $omp{'end'}   = $4 || $2 || 1;
}

mkdir "report" if ( !-d "report" );

my $command;
my $timers;
my $total = @model * ( $mpi{'end'} - $mpi{'begin'} + 1 ) * ( $omp{'end'} - $omp{'begin'} + 1 ) * 3;
my $current = 0;
my $variant;
my $new_variant;
foreach my $var ( @model ) {
    open($variant,'<Dinar.dat');
    open($new_variant,'>new_Dinar.dat');
    select($new_variant);
    while ( <$variant> ) {
        s/NHARM  \d+/NHARM  $var/;
        print;
    }
    close($variant);
    close($new_variant);
    unlink 'Dinar.dat';
    move('new_Dinar.dat','Dinar.dat');
    select(STDOUT);
    for ( my $i = $mpi{'begin'}; $i <= $mpi{'end'}; $i++ ) {
        for ( my $j = $omp{'begin'}; $j <= $omp{'end'}; $j++ ) {
            $command = ( $mpi{'using'} ? "mpirun -np $i " : '' ).'bars_ktp6.exe'.( $omp{'using'} ? " $j" : '' );
            for ( my $k = 1; $k <= 3; $k++ ) {
                $timers = "timers$var\_$i\_$j\_$k.txt";
                $current += 1;
                print "$command ........";
                $ENV{OMP_NESTED} = 1;
                $ENV{OMP_DYNAMIC} = 0;
                $ENV{OMP_NUM_THREADS} = $j;
                `$command` and print "ok\n";
                move('timers.txt',"report/$timers") and print "$timers ........".( $current / $total * 100 )."%\n" or die;
            }
        }
    }
}