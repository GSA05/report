use common::sense;
use strict;
use Data::Dumper;
use Excel::Writer::XLSX;
use POSIX qw(ceil);
use Math::MatrixReal;
#use re "debug";

my $thread_num = 8;
my $student = 1.3862;#2.92;
my $NG = 5;
#my $OMP = 16;
my $MPI = 1;

# Заголовки
my @names = (
    'Идеал',
    'Закон Амдала',
    'Модифицированный Закон Амдала',
    'Эксперимент',
);

my $i = 0;
my $j = 0;
my $k = 0;
my %time = undef;
my $file;

my %functions = (
    CONST => {
    },
    INIT => {
        SIZE => 'NH',
        },
    SOURCE => {
        SIZE => 'NH',
        },
    GMATR => {
        SIZE => 'NG*NG',
        },
    GMATRF => {
        SIZE => 'NG*NG',
        },
    GMATRL => {
        SIZE => 'NG*NG',
        },
    GMATR_F => {
        SIZE => 'NG*NG',
        },
    GMATR_C => {
        SIZE => 'NG*NG',
        },
    INTEG => {
        SIZE => 'NH',
        },
    INTEGL => {
        SIZE => 'NH',
        },
    );

sub sum {
    my ($data,$pow) = @_;
    $pow = $pow || 1;
    my $s = 0;        
    foreach ( @$data ) { $s += $_**$pow }
    return $s;
}

sub mean {
    my $data = shift;
    my $n = @$data;
    my $m = sum($data);
    $m /= $n;
    return $m;
}

sub variance {
    my $data = shift;
    my $n = @$data;
    my $s = sum($data);
    my $v = sum($data,2);
    $v = ($v - $s**2/$n)/($n-1);
    return $v;
}

sub am_var {
    my ( $means, $vars, $N , $t ) = @_;
    my $vars_vec = Math::MatrixReal->new_from_rows([$vars]);
    my $means_vec = Math::MatrixReal->new_from_rows([$means]);
    my ( undef, $size ) = $means_vec->dim();
    $vars_vec = $vars_vec->each( sub {
        my ( $val, $i, $j ) = @_;
        $val /= $N->[$j - 1];
    } );
    my $C = new Math::MatrixReal($size, $size);
    for ( my $i = 1; $i <= $size; $i++) { $C = $C->assign_row($i, $means_vec) }
    #print Dumper($N);
    $C = $C->each( sub {
        my ( $val, $i, $j ) = @_;
        $val = $val * ($N->[$i - 1] - $N->[$j - 1]) / $N->[$j - 1];
    } );
    print Dumper($C) if $t;
    return ($vars_vec * $C * ~$C * ~$vars_vec)->element(1, 1);
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

#print Dumper(%functions);

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

foreach $key ( sort keys %time ) {
    my $f = 0;
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
                if ( defined $functions{$name} ) {
                    $functions{$name}{'POSITION'} = $k * 1 + $col - 1;
                    my $size = $functions{$name}{'SIZE'};
                    $size =~ s/NG/$NG/g;
                    $size =~ s/NH/$key/g;
                    $functions{$name}{'SIZES'}{$key} = eval $size;
                    $functions{'CONST'}{$key}{'PROC_MEAN'}[$f] = $time{$key}->[0]{$name}{'PROC_MEAN'};
                    $functions{'CONST'}{$key}{$name}{'PROC_MEAN'} = $time{$key}->[0]{$name}{'PROC_MEAN'};
                    $functions{'CONST'}{$key}{'PROC_VAR'}[$f++] = $student * $time{$key}->[0]{$name}{'PROC_VAR'}**0.5;
                    $functions{'CONST'}{$key}{$name}{'PROC_VAR'} = $student * $time{$key}->[0]{$name}{'PROC_VAR'}**0.5;
                }
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
    $col = 0;
    $row++;
    my $c = 0;#$functions{'CONST'}{$key}{'PROC'} = 100 - sum($functions{'CONST'}{$key}{'PROC_MEAN'});
    my $c_var = $functions{'CONST'}{$key}{'VAR'} = sum($functions{'CONST'}{$key}{'PROC_VAR'},2)**0.5;
    my $p = sum($functions{'CONST'}{$key}{'PROC_MEAN'});
    my $p_var = $c_var;
    $worksheet->write($row,$col++,'OMP',$format);
    $worksheet->write($row,$col++,scalar @{$time{$key}},$format2);
    $worksheet->write($row,$col++,'MPI',$format);
    $worksheet->write($row,$col++,$MPI,$format2);
    $worksheet->write($row,$col++,'NG',$format);
    $worksheet->write($row,$col++,$NG,$format2);
    $worksheet->write($row,$col++,'NH',$format);
    $worksheet->write($row,$col++,$key,$format2);
    $worksheet->write($row,$col++,'C',$format);
    $worksheet->write($row,$col++,$c,$format2);
    $worksheet->write($row,$col++,'C_VAR',$format);
    $worksheet->write($row,$col++,$c_var,$format2);
    $col = 0;
    $row++;
    $worksheet->write($row,$col++,'N',$format);
    $worksheet->write($row,$col++,'A',$format);
    $worksheet->write($row,$col++,'A_VAR',$format);
    $worksheet->write($row,$col++,'AM',$format);
    $worksheet->write($row,$col++,'AM_VAR',$format);
    $worksheet->write($row,$col++,'RES',$format);
    $worksheet->write($row,$col++,'RES_VAR',$format);
    $i = 1;
    $row++;
    my $a;
    my $am = 0;
    my $pm = 0;
    my ( @p_means, @p_vars, @p_N );
    push @p_means, $c;
    push @p_vars, $c_var;
    push @p_N, 1;
    while ( $i <= @{$time{$key}} ) {
        $a = ($c + $p)/($c + $p/$i);
        $am = 0;
        @p_N = ( 1 );
        my $o = $time{$key}->[0]{'ALL'}{'MEAN'}*(1-$p/100);
        $mean0 = $time{$key}->[0]{'ALL'}{'MEAN'} - $o;
        $var0 = $student * $time{$key}->[0]{'ALL'}{'VAR'}**0.5;
        $mean = $time{$key}->[$i - 1]{'ALL'}{'MEAN'} - $o;
        $var = $student * $time{$key}->[$i - 1]{'ALL'}{'VAR'}**0.5;
        foreach my $name ( keys %functions ) {
            next if $name eq 'CONST';
            my $pos = $functions{$name}{'POSITION'};
            my $N = $functions{$name}{'SIZES'}{$key};
            $N = $N/ceil($N/$i);
            push @p_N, $N;
            $am += $functions{'CONST'}{$key}{$name}{'PROC_MEAN'}/$N;
            if ( $i == 1 ) {
                $pm += $functions{'CONST'}{$key}{$name}{'PROC_MEAN'};
                push @p_means, $functions{'CONST'}{$key}{$name}{'PROC_MEAN'};
                push @p_vars, $functions{'CONST'}{$key}{$name}{'PROC_VAR'};
                $worksheet->write($row - 1,$pos,'NM',$format);
            }
            $worksheet->write($row,$pos,$N,$format2);
        }
        $am = ($c + $pm)/($c + $am);
        $worksheet->write($row,0,$i);
        $worksheet->write($row,1,$a,$format2);
        #$worksheet->write($row,2,$a**2/100 * ($c_var**2 + $p_var**2/$i**2)**0.5,$format2);
        $worksheet->write($row,2,($i - 1)/$i * $a**2/($c + $p)**2 * ($c_var**2 * $p**2 + $p_var**2 * $c**2)**0.5,$format2);
        $worksheet->write($row,3,$am,$format2);
        #print Dumper(@p_vars);
        $worksheet->write($row,4,$am**2/($c + $pm)**2 * (am_var(\@p_means, \@p_vars, \@p_N, ( $i == 0 && $key == 20 )))**0.5,$format2);
        $worksheet->write($row,5,$mean0/$mean,$format2);
        $worksheet->write($row,6,($var0**2 * $mean**2 + $mean0**2 * $var**2)**0.5/$mean**2,$format2);
        $row++;
        $i++;
    }
    my $chart = $workbook->add_chart(
        type => 'line',
        embedded => 1,
    );
    $chart->add_series(
        name       => @names[0],
        categories => "=$key!A20:A35",
        values     => "=$key!A20:A35",
    );
    $chart->add_series(
        name         => @names[1],
        categories   => "=$key!A20:A35",
        values       => "=$key!B20:B35",
        y_error_bars => {
            type         => 'custom',
            plus_values  => "=$key!C20:C35",
            minus_values => "=$key!C20:C35",
        },
    );
    $chart->add_series(
        name       => @names[2],
        categories => "=$key!A20:A35",
        values     => "=$key!D20:D35",
        y_error_bars => {
            type         => 'custom',
            plus_values  => "=$key!E20:E35",
            minus_values => "=$key!E20:E35",
        },
    );
    $chart->add_series(
        name       => @names[3],
        categories => "=$key!A20:A35",
        values     => "=$key!F20:F35",
        y_error_bars => {
            type         => 'custom',
            plus_values  => "=$key!G20:G35",
            minus_values => "=$key!G20:G35",
        },
    );
    $chart->set_title ( name => 'Результат распараллеливания' );
    $chart->set_x_axis( name => 'Количество ядер' );
    $chart->set_y_axis( name => 'Коэффициент распараллеливания' );
    $chart->set_style( 2 );
    $worksheet->insert_chart('H19',$chart,30,15);
}

#print Dumper(%functions);