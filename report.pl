use strict;
use Data::Dumper;
use Excel::Writer::XLSX;
#use re "debug";

my $thread_num = 8;
my $student = 1.3862;

my $i = 0;
my $j = 0;
my $k = 0;
my %time = undef;
my $file;

sub mean {
    my $data = shift;
    my $n = 3;
    my $m = 0;
        
    for ( my $i = 0; $i < $n; $i++ ) { $m += $data->[$i] }
    $m /= $n;
    return $m;
}

sub variance {
    my $data = shift;
    my $n = 3;
    my $s = 0;
    my $v = 0;
    
    for ( my $i = 0; $i < $n; $i++ ) {
        $v += $data->[$i]**2;
        $s += $data->[$i];
    }
    $v = ($v - $s**2/$n)/($n-1);
    return $v;
}

while ( <*.txt> ) {
    if ( /timers(.+)_(\d+)_(\d+)_(\d+)\.txt/ ) {
        $i = $1;
        $j = $3;
        $k = $4;
        open FILE, $_;
        while ( <FILE> ) {
            if ( /[^A-Z]+([A-Z_]+):\D*(\d+\.\d+)(\D*(\d+\.\d+))?/ ) {
                $time{$i}[$j-1]{$1}{'DATA'}[$k-1] = "$2";
                $time{$i}[$j-1]{$1}{'PROC'}[$k-1] = "$4";
            }
        }
        close FILE;
    }
}

#print Dumper(%time);

delete $time{""};

#print Dumper(%time);

for $i ( values %time ) {
    for ( $j = 0; $j < @$i; $j++ ) {
        for $k ( keys $i->[$j] ) {
            $i->[$j]->{$k}{'MEAN'} = mean($i->[$j]->{$k}{'DATA'});
            $i->[$j]->{$k}{'VAR'}  = variance($i->[$j]->{$k}{'DATA'});
            $i->[$j]->{$k}{'PROC_MEAN'} = mean($i->[$j]->{$k}{'PROC'});
            $i->[$j]->{$k}{'PROC_VAR'}  = variance($i->[$j]->{$k}{'PROC'});
        }
    }
}

my $workbook = Excel::Writer::XLSX->new('report.xlsx');
my $format = $workbook->add_format();
$format->set_bold();
$format->set_color('red');
$format->set_align('center');
my $format2 = $workbook->add_format();
$format2->set_align('left');
$format2->set_num_format('0.00');

my $col;
my $row;
my $key;
my $colc = 10;
$k = 0;
my $mean0;
my $var0;
my $mean;
my $var;
my %positions;
foreach $key ( keys %time ) {
    my $worksheet = $workbook->add_worksheet($key);
    for ( $row = 0; $row < @{$time{$key}}; $row++) {
        $col = 0;
        $k = 0;
        foreach my $name ( sort keys %{$time{$key}->[$row]} ) {
            $mean0 = $time{$key}->[0]{$name}{'MEAN'};
            $var0 = $student * $time{$key}->[0]{$name}{'VAR'}**0.5;
            $mean = $time{$key}->[$row]{$name}{'MEAN'};
            $var = $student * $time{$key}->[$row]{$name}{'VAR'}**0.5;
            if ( !$row ) {
                $worksheet->write($row,$k * 1 + $col++,$name,$format);
                if ( $name eq 'INIT'
                     or $name eq 'SOURCE'
                     or $name eq 'GMATR'
                     or $name eq 'GMATRF'
                     or $name eq 'GMATRL'
                     or $name eq 'GMATR_F'
                     or $name eq 'GMATR_C'
                     or $name eq 'INTEG'
                     or $name eq 'INTEGL' ) { $positions{$name} = $col - 1 }
                $worksheet->write($row,$k * 1 + $col++,'DATA1',$format);
                $worksheet->write($row,$k * 1 + $col++,'DATA2',$format);
                $worksheet->write($row,$k * 1 + $col++,'DATA3',$format);
                $worksheet->write($row,$k * 1 + $col++,'MEAN',$format);
                $worksheet->write($row,$k * 1 + $col++,'VAR',$format);
                $worksheet->write($row,$k * 1 + $col++,'COEF',$format);
                $worksheet->write($row,$k * 1 + $col++,'COEF_VAR',$format);
                $worksheet->write($row,$k * 1 + $col++,'EFFI',$format);
                $col -= $colc - 1;
            }
            #else { $col++ }
            $worksheet->write($row + 1,$k * 1 + $col++,$row + 1);
            for ( $i = 0; $i < @{$time{$key}->[$row]{$name}{'DATA'}}; $i++ ) {
                $worksheet->write($row + 1,$k * 1 + $col++,$time{$key}->[$row]{$name}{'DATA'}[$i],$format2);
            }
            $worksheet->write($row + 1,$k * 1 + $col++,$mean,$format2);
            $worksheet->write($row + 1,$k * 1 + $col++,$var,$format2);
            if ( $mean > 0.001 ) { $worksheet->write($row + 1,$k * 1 + $col++,$mean0/$mean,$format2) }
            else { $worksheet->write($row + 1,$k * 1 + $col++,'0') }
            if ( $mean > 0.001 ) { $worksheet->write($row + 1,$k * 1 + $col++,(($var0 * $mean)**2 + ($mean0 * $var)**2)**0.5/($mean)**2,$format2) }
            else { $worksheet->write($row + 1,$k * 1 + $col++,'0') }
            if ( $mean > 0.001 ) { $worksheet->write($row + 1,$k * 1 + $col++,$mean0/$mean/($row + 1),$format2) }
            else { $worksheet->write($row + 1,$k * 1 + $col++,'0') }
            $k++;
        }
    }
    
}

