# alte re:
# {^(\d+)\.(\d+)\.(\d+)T?\s*(\d+)?:?(\d+)?:?(\d+)?\.?(\d+)?\s*([+-])?(\d+)?:?(\d+)?$} $scan all d m y H M S F x a b
# {^(\d+)-(\d+)-(\d+)T?\s*(\d+)?:?(\d+)?:?(\d+)?\.?(\d+)?\s*([+-])?(\d+)?:?(\d+)?$}   $scan all y m d H M S F x a b
# {^(\d+)-(\w+)-(\d+)T?\s*(\d+)?:?(\d+)?:?(\d+)?\.?(\d+)?\s*([+-])?(\d+)?:?(\d+)?$}   $scan all d ml y H M S F x a b
# {^(\d+)/(\d+)/(\d+)T?\s*(\d+)?:?(\d+)?:?(\d+)?\.?(\d+)?\s*([+-])?(\d+)?:?(\d+)?$}   $scan all m d y H M S F x a b

set strings {
    {2025-03-27}                          {date only}
    {2025-03-27-10}                       {too many separators}
    {2025-03-27T14:05:30}                 {with time (T)}
    {2025-03-27 14:05:30}                 {with time (space)}
    {2025-03-27T14:05:30.789}             {with milliseconds (T)}
    {2025-03-27 14:05:30.789}             {with milliseconds (space)}
    {2025-03-27T14:05:30Z}                {with UTC timezone (T)}
    {2025-03-27 14:05:30Z}                {with UTC timezone (space)}
    {2025-03-27T14:05:30.789Z}            {with milliseconds and UTC (T)}
    {2025-03-27 14:05:30.789Z}            {with milliseconds and UTC (space)}
    {2025-03-27T14:05:30+02:15}           {with positive timezone (T)}
    {2025-03-27 14:05:30+02:15}           {with positive timezone (space)}
    {2025-03-27T14:05:30-02:15}           {with negative timezone (T)}
    {2025-03-27 14:05:30-02:15}           {with negative timezone (space)}
    {2025-03-27T14:05:30.789+02:15}       {with milliseconds & timezone (T)}
    {2025-03-27 14:05:30.789+02:15}       {with milliseconds & timezone (space)}
    {2025-3-27T14:05}                     {month not two-digit}
    {2025-03-27T14:5:30}                  {minute not two-digit}
    {2025-03-27T14:05:30.}                {dot without fractional part}
    {2025-03-27T14:05:30+2:15}            {hour part of timezone not two-digit}
    {2025-03-27T14:05:30+0215}            {timezone without colon}
    {2025-03-27 14:05}                    {seconds missing}
    {2025-03-27 14:05:}                   {colon and seconds missing}
    {2025-03-27T14:05:30Z+02:15}          {two timezone specifications simultaneously}
    {2025-03-27T}                         {only "T" without time}
    {2025-03-27 14}                       {incomplete time specification}
    {2025-03-27T14:05:30-}                {only minus without timezone}
    {27.03.2025}                          {date only}
    {27.03.2025.10}                       {too many separators}
    {27.03.2025T14:05:30}                 {with time (T)}
    {27.03.2025 14:05:30}                 {with time (space)}
    {27.03.2025T14:05:30.789}             {with milliseconds (T)}
    {27.03.2025 14:05:30.789}             {with milliseconds (space)}
    {27.03.2025T14:05:30Z}                {with UTC timezone (T)}
    {27.03.2025 14:05:30Z}                {with UTC timezone (space)}
    {27.03.2025T14:05:30.789Z}            {with milliseconds and UTC (T)}
    {27.03.2025 14:05:30.789Z}            {with milliseconds and UTC (space)}
    {27.03.2025T14:05:30+02:15}           {with positive timezone (T)}
    {27.03.2025 14:05:30+02:15}           {with positive timezone (space)}
    {27.03.2025T14:05:30-02:15}           {with negative timezone (T)}
    {27.03.2025 14:05:30-02:15}           {with negative timezone (space)}
    {27.03.2025T14:05:30.789+02:15}       {with milliseconds & timezone (T)}
    {27.03.2025 14:05:30.789+02:15}       {with milliseconds & timezone (space)}
    {27.3.2025 14:05:30}                  {month not two-digit}
    {27.03.2025T14:5:30}                  {minute not two-digit}
    {27.03.2025T14:05:30.}                {dot without fractional part}
    {27.03.2025T14:05:30+2:15}            {hour part of timezone not two-digit}
    {27.03.2025T14:05:30+0215}            {timezone without colon}
    {27.03.2025 14:05:}                   {seconds missing}
    {27.03.2025T14:05:30Z+02:15}          {two timezone specifications simultaneously}
    {27.03.2025T}                         {only "T" without time}
    {27.03.2025 14}                       {incomplete time specification}
    {27.03.2025T14:05:30-}                {only minus without timezone}
    {27-mar-2025}                         {date only}
    {27-mar-2025-10}                      {too many separators}
    {27-mar-2025T14:05:30}                {with time (T)}
    {27-mar-2025 14:05:30}                {with time (space)}
    {27-mar-2025T14:05:30.789}            {with milliseconds (T)}
    {27-mar-2025 14:05:30.789}            {with milliseconds (space)}
    {27-mar-2025T14:05:30Z}               {with UTC timezone (T)}
    {27-mar-2025 14:05:30Z}               {with UTC timezone (space)}
    {27-mar-2025T14:05:30.789Z}           {with milliseconds and UTC (T)}
    {27-mar-2025 14:05:30.789Z}           {with milliseconds and UTC (space)}
    {27-mar-2025T14:05:30+02:15}          {with positive timezone (T)}
    {27-mar-2025 14:05:30+02:15}          {with positive timezone (space)}
    {27-mar-2025T14:05:30-02:15}          {with negative timezone (T)}
    {27-mar-2025 14:05:30-02:15}          {with negative timezone (space)}
    {27-mar-2025T14:05:30.789+02:15}      {with milliseconds & timezone (T)}
    {27-mar-2025 14:05:30.789+02:15}      {with milliseconds & timezone (space)}
    {27-mar-2025T14:5:30}                 {minute not two-digit}
    {27-mar-2025T14:05:30.}               {dot without fractional part}
    {27-mar-2025T14:05:30+2:15}           {hour part of timezone not two-digit}
    {27-mar-2025T14:05:30+0215}           {timezone without colon}
    {27-mar-2025 14:05:}                  {seconds missing}
    {27-mar-2025T14:05:30Z+02:15}         {two timezone specifications simultaneously}
    {27-mar-2025T}                        {only "T" without time}
    {27-mar-2025 14}                      {incomplete time specification}
    {27-mar-2025T14:05:30-}               {only minus without timezone}
    {03/27/2025}                          {date only}
    {03/27/2025/10}                       {too many separators}
    {03/27/2025T14:05:30}                 {with time (T)}
    {03/27/2025 14:05:30}                 {with time (space)}
    {03/27/2025T14:05:30.789}             {with milliseconds (T)}
    {03/27/2025 14:05:30.789}             {with milliseconds (space)}
    {03/27/2025T14:05:30Z}                {with UTC timezone (T)}
    {03/27/2025 14:05:30Z}                {with UTC timezone (space)}
    {03/27/2025T14:05:30.789Z}            {with milliseconds and UTC (T)}
    {03/27/2025 14:05:30.789Z}            {with milliseconds and UTC (space)}
    {03/27/2025T14:05:30+02:15}           {with positive timezone (T)}
    {03/27/2025 14:05:30+02:15}           {with positive timezone (space)}
    {03/27/2025T14:05:30-02:15}           {with negative timezone (T)}
    {03/27/2025 14:05:30-02:15}           {with negative timezone (space)}
    {03/27/2025T14:05:30.789+02:15}       {with milliseconds & timezone (T)}
    {03/27/2025 14:05:30.789+02:15}       {with milliseconds & timezone (space)}
    {3/27/2025T14:05:30}                  {month not two-digit}
    {03/27/2025T14:5:30}                  {minute not two-digit}
    {03/27/2025T14:05:30.}                {dot without fractional part}
    {03/27/2025T14:05:30+2:15}            {hour part of timezone not two-digit}
    {03/27/2025T14:05:30+0215}            {timezone without colon}
    {03/27/2025 14:05:}                   {seconds missing}
    {03/27/2025T14:05:30Z+02:15}          {two timezone specifications simultaneously}
    {03/27/2025T}                         {only "T" without time}
    {03/27/2025 14}                       {incomplete time specification}
    {03/27/2025T14:05:30-}                {only minus without timezone}
}

set re {^(\d{2,4})-(\d{2})-(\d{2})(?:[ T]|$)(?:(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?(?:([+-])(\d{2}):(\d{2})|(Z)$)?)?)?$}
puts $re
foreach {toScan comment} $strings {
    if {[regexp $re $toScan all y m d H M S F x a b z]} {
        puts [format "%-32s %-46s 1-all  %-30s  %-4s  %-2s  %-2s  %-2s  %-2s  %-2s  %-3s %-1s  %-2s  %-2s  %s" $toScan $comment $all $y $m $d $H $M $S $F $x $a $b $z]
    } else {
        puts [format "%-32s %-46s 1-FAILED" $toScan $comment]
    }
}

set re {^(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:[ T]|$)(?:(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?(?:([+-])(\d{1,2}):(\d{1,2})|(Z))?)?)?$}
puts $re
foreach {toScan comment} $strings {
    if {[regexp $re $toScan all d m y H M S F x a b z]} {
        puts [format "%-32s %-46s 2-all  %-30s  %-4s  %-2s  %-2s  %-2s  %-2s  %-2s  %-3s %-1s  %-2s  %-2s  %s" $toScan $comment $all $y $m $d $H $M $S $F $x $a $b $z]
    } else {
        puts [format "%-32s %-46s 2-FAILED" $toScan $comment]
    }
}

set re {^(\d{1,2})-(\w+)-(\d{2,4})(?:[ T]|$)(?:(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?(?:([+-])(\d{1,2}):(\d{1,2})|(Z))?)?)?$}
puts $re
foreach {toScan comment} $strings {
    if {[regexp $re $toScan all d m y H M S F x a b z]} {
        puts [format "%-32s %-46s 3-all  %-30s  %-4s  %-2s  %-2s  %-2s  %-2s  %-2s  %-3s %-1s  %-2s  %-2s  %s" $toScan $comment $all $y $m $d $H $M $S $F $x $a $b $z]
    } else {
        puts [format "%-32s %-46s 3-FAILED" $toScan $comment]
    }
}

set re {^(\d{1,2})/(\d{1,2})/(\d{2,4})(?:[ T]|$)(?:(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?(?:([+-])(\d{1,2}):(\d{1,2})|(Z))?)?)?$}
puts $re
foreach {toScan comment} $strings {
    if {[regexp $re $toScan all m d y H M S F x a b z]} {
        puts [format "%-32s %-46s 4-all  %-30s  %-4s  %-2s  %-2s  %-2s  %-2s  %-2s  %-3s %-1s  %-2s  %-2s  %s" $toScan $comment $all $y $m $d $H $M $S $F $x $a $b $z]
    } else {
        puts [format "%-32s %-46s 4-FAILED" $toScan $comment]
    }
}
