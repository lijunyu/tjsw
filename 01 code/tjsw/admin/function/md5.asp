<% 
private const bits_to_a_byte = 8 
private const bytes_to_a_word = 4 
private const bits_to_a_word = 32 
private m_lonbits(30) 
private m_l2power(30) 

private function lshift(lvalue, ishiftbits) 
if ishiftbits = 0 then 
lshift = lvalue 
exit function 
elseif ishiftbits = 31 then 
if lvalue and 1 then 
lshift = &h80000000 
else 
lshift = 0 
end if 
exit function 
elseif ishiftbits < 0 or ishiftbits > 31 then 
err.raise 6 
end if 

if (lvalue and m_l2power(31 - ishiftbits)) then 
lshift = ((lvalue and m_lonbits(31 - (ishiftbits + 1))) * m_l2power(ishiftbits)) or &h80000000 
else 
lshift = ((lvalue and m_lonbits(31 - ishiftbits)) * m_l2power(ishiftbits)) 
end if 
end function 

private function rshift(lvalue, ishiftbits) 
if ishiftbits = 0 then 
rshift = lvalue 
exit function 
elseif ishiftbits = 31 then 
if lvalue and &h80000000 then 
rshift = 1 
else 
rshift = 0 
end if 
exit function 
elseif ishiftbits < 0 or ishiftbits > 31 then 
err.raise 6 
end if 

rshift = (lvalue and &h7ffffffe) \ m_l2power(ishiftbits) 

if (lvalue and &h80000000) then 
rshift = (rshift or (&h40000000 \ m_l2power(ishiftbits - 1))) 
end if 
end function 

private function rotateleft(lvalue, ishiftbits) 
rotateleft = lshift(lvalue, ishiftbits) or rshift(lvalue, (32 - ishiftbits)) 
end function 

private function addunsigned(lx, ly) 
dim lx4 
dim ly4 
dim lx8 
dim ly8 
dim lresult 

lx8 = lx and &h80000000 
ly8 = ly and &h80000000 
lx4 = lx and &h40000000 
ly4 = ly and &h40000000 

lresult = (lx and &h3fffffff) + (ly and &h3fffffff) 

if lx4 and ly4 then 
lresult = lresult xor &h80000000 xor lx8 xor ly8 
elseif lx4 or ly4 then 
if lresult and &h40000000 then 
lresult = lresult xor &hc0000000 xor lx8 xor ly8 
else 
lresult = lresult xor &h40000000 xor lx8 xor ly8 
end if 
else 
lresult = lresult xor lx8 xor ly8 
end if 

addunsigned = lresult 
end function 

private function md5_f(x, y, z) 
md5_f = (x and y) or ((not x) and z) 
end function 

private function md5_g(x, y, z) 
md5_g = (x and z) or (y and (not z)) 
end function 

private function md5_h(x, y, z) 
md5_h = (x xor y xor z) 
end function 

private function md5_i(x, y, z) 
md5_i = (y xor (x or (not z))) 
end function 

private sub md5_ff(a, b, c, d, x, s, ac) 
a = addunsigned(a, addunsigned(addunsigned(md5_f(b, c, d), x), ac)) 
a = rotateleft(a, s) 
a = addunsigned(a, b) 
end sub 

private sub md5_gg(a, b, c, d, x, s, ac) 
a = addunsigned(a, addunsigned(addunsigned(md5_g(b, c, d), x), ac)) 
a = rotateleft(a, s) 
a = addunsigned(a, b) 
end sub 

private sub md5_hh(a, b, c, d, x, s, ac) 
a = addunsigned(a, addunsigned(addunsigned(md5_h(b, c, d), x), ac)) 
a = rotateleft(a, s) 
a = addunsigned(a, b) 
end sub 

private sub md5_ii(a, b, c, d, x, s, ac) 
a = addunsigned(a, addunsigned(addunsigned(md5_i(b, c, d), x), ac)) 
a = rotateleft(a, s) 
a = addunsigned(a, b) 
end sub 

private function converttowordarray(smessage) 
dim lmessagelength 
dim lnumberofwords 
dim lwordarray() 
dim lbyteposition 
dim lbytecount 
dim lwordcount 

const modulus_bits = 512 
const congruent_bits = 448 

lmessagelength = len(smessage) 

lnumberofwords = (((lmessagelength + ((modulus_bits - congruent_bits) \ bits_to_a_byte)) \ (modulus_bits \ bits_to_a_byte)) + 1) * (modulus_bits \ bits_to_a_word) 
redim lwordarray(lnumberofwords - 1) 

lbyteposition = 0 
lbytecount = 0 
do until lbytecount >= lmessagelength 
lwordcount = lbytecount \ bytes_to_a_word 
lbyteposition = (lbytecount mod bytes_to_a_word) * bits_to_a_byte 
lwordarray(lwordcount) = lwordarray(lwordcount) or lshift(asc(mid(smessage, lbytecount + 1, 1)), lbyteposition) 
lbytecount = lbytecount + 1 
loop 

lwordcount = lbytecount \ bytes_to_a_word 
lbyteposition = (lbytecount mod bytes_to_a_word) * bits_to_a_byte 

lwordarray(lwordcount) = lwordarray(lwordcount) or lshift(&h80, lbyteposition) 

lwordarray(lnumberofwords - 2) = lshift(lmessagelength, 3) 
lwordarray(lnumberofwords - 1) = rshift(lmessagelength, 29) 

converttowordarray = lwordarray 
end function 

private function wordtohex(lvalue) 
dim lbyte 
dim lcount 

for lcount = 0 to 3 
lbyte = rshift(lvalue, lcount * bits_to_a_byte) and m_lonbits(bits_to_a_byte - 1) 
wordtohex = wordtohex & right("0" & hex(lbyte), 2) 
next 
end function 

public function md5(smessage) 
m_lonbits(0) = clng(1) 
m_lonbits(1) = clng(3) 
m_lonbits(2) = clng(7) 
m_lonbits(3) = clng(15) 
m_lonbits(4) = clng(31) 
m_lonbits(5) = clng(63) 
m_lonbits(6) = clng(127) 
m_lonbits(7) = clng(255) 
m_lonbits(8) = clng(511) 
m_lonbits(9) = clng(1023) 
m_lonbits(10) = clng(2047) 
m_lonbits(11) = clng(4095) 
m_lonbits(12) = clng(8191) 
m_lonbits(13) = clng(16383) 
m_lonbits(14) = clng(32767) 
m_lonbits(15) = clng(65535) 
m_lonbits(16) = clng(131071) 
m_lonbits(17) = clng(262143) 
m_lonbits(18) = clng(524287) 
m_lonbits(19) = clng(1048575) 
m_lonbits(20) = clng(2097151) 
m_lonbits(21) = clng(4194303) 
m_lonbits(22) = clng(8388607) 
m_lonbits(23) = clng(16777215) 
m_lonbits(24) = clng(33554431) 
m_lonbits(25) = clng(67108863) 
m_lonbits(26) = clng(134217727) 
m_lonbits(27) = clng(268435455) 
m_lonbits(28) = clng(536870911) 
m_lonbits(29) = clng(1073741823) 
m_lonbits(30) = clng(2147483647) 

m_l2power(0) = clng(1) 
m_l2power(1) = clng(2) 
m_l2power(2) = clng(4) 
m_l2power(3) = clng(8) 
m_l2power(4) = clng(16) 
m_l2power(5) = clng(32) 
m_l2power(6) = clng(64) 
m_l2power(7) = clng(128) 
m_l2power(8) = clng(256) 
m_l2power(9) = clng(512) 
m_l2power(10) = clng(1024) 
m_l2power(11) = clng(2048) 
m_l2power(12) = clng(4096) 
m_l2power(13) = clng(8192) 
m_l2power(14) = clng(16384) 
m_l2power(15) = clng(32768) 
m_l2power(16) = clng(65536) 
m_l2power(17) = clng(131072) 
m_l2power(18) = clng(262144) 
m_l2power(19) = clng(524288) 
m_l2power(20) = clng(1048576) 
m_l2power(21) = clng(2097152) 
m_l2power(22) = clng(4194304) 
m_l2power(23) = clng(8388608) 
m_l2power(24) = clng(16777216) 
m_l2power(25) = clng(33554432) 
m_l2power(26) = clng(67108864) 
m_l2power(27) = clng(134217728) 
m_l2power(28) = clng(268435456) 
m_l2power(29) = clng(536870912) 
m_l2power(30) = clng(1073741824) 

dim x 
dim k 
dim aa 
dim bb 
dim cc 
dim dd 
dim a 
dim b 
dim c 
dim d 

const s11 = 7 
const s12 = 12 
const s13 = 17 
const s14 = 22 
const s21 = 5 
const s22 = 9 
const s23 = 14 
const s24 = 20 
const s31 = 4 
const s32 = 11 
const s33 = 16 
const s34 = 23 
const s41 = 6 
const s42 = 10 
const s43 = 15 
const s44 = 21 

x = converttowordarray(smessage) 

a = &h67452301 
b = &hefcdab89 
c = &h98badcfe 
d = &h10325476 

for k = 0 to ubound(x) step 16 
aa = a 
bb = b 
cc = c 
dd = d 

md5_ff a, b, c, d, x(k + 0), s11, &hd76aa478 
md5_ff d, a, b, c, x(k + 1), s12, &he8c7b756 
md5_ff c, d, a, b, x(k + 2), s13, &h242070db 
md5_ff b, c, d, a, x(k + 3), s14, &hc1bdceee 
md5_ff a, b, c, d, x(k + 4), s11, &hf57c0faf 
md5_ff d, a, b, c, x(k + 5), s12, &h4787c62a 
md5_ff c, d, a, b, x(k + 6), s13, &ha8304613 
md5_ff b, c, d, a, x(k + 7), s14, &hfd469501 
md5_ff a, b, c, d, x(k + 8), s11, &h698098d8 
md5_ff d, a, b, c, x(k + 9), s12, &h8b44f7af 
md5_ff c, d, a, b, x(k + 10), s13, &hffff5bb1 
md5_ff b, c, d, a, x(k + 11), s14, &h895cd7be 
md5_ff a, b, c, d, x(k + 12), s11, &h6b901122 
md5_ff d, a, b, c, x(k + 13), s12, &hfd987193 
md5_ff c, d, a, b, x(k + 14), s13, &ha679438e 
md5_ff b, c, d, a, x(k + 15), s14, &h49b40821 

md5_gg a, b, c, d, x(k + 1), s21, &hf61e2562 
md5_gg d, a, b, c, x(k + 6), s22, &hc040b340 
md5_gg c, d, a, b, x(k + 11), s23, &h265e5a51 
md5_gg b, c, d, a, x(k + 0), s24, &he9b6c7aa 
md5_gg a, b, c, d, x(k + 5), s21, &hd62f105d 
md5_gg d, a, b, c, x(k + 10), s22, &h2441453 
md5_gg c, d, a, b, x(k + 15), s23, &hd8a1e681 
md5_gg b, c, d, a, x(k + 4), s24, &he7d3fbc8 
md5_gg a, b, c, d, x(k + 9), s21, &h21e1cde6 
md5_gg d, a, b, c, x(k + 14), s22, &hc33707d6 
md5_gg c, d, a, b, x(k + 3), s23, &hf4d50d87 
md5_gg b, c, d, a, x(k + 8), s24, &h455a14ed 
md5_gg a, b, c, d, x(k + 13), s21, &ha9e3e905 
md5_gg d, a, b, c, x(k + 2), s22, &hfcefa3f8 
md5_gg c, d, a, b, x(k + 7), s23, &h676f02d9 
md5_gg b, c, d, a, x(k + 12), s24, &h8d2a4c8a 

md5_hh a, b, c, d, x(k + 5), s31, &hfffa3942 
md5_hh d, a, b, c, x(k + 8), s32, &h8771f681 
md5_hh c, d, a, b, x(k + 11), s33, &h6d9d6122 
md5_hh b, c, d, a, x(k + 14), s34, &hfde5380c 
md5_hh a, b, c, d, x(k + 1), s31, &ha4beea44 
md5_hh d, a, b, c, x(k + 4), s32, &h4bdecfa9 
md5_hh c, d, a, b, x(k + 7), s33, &hf6bb4b60 
md5_hh b, c, d, a, x(k + 10), s34, &hbebfbc70 
md5_hh a, b, c, d, x(k + 13), s31, &h289b7ec6 
md5_hh d, a, b, c, x(k + 0), s32, &heaa127fa 
md5_hh c, d, a, b, x(k + 3), s33, &hd4ef3085 
md5_hh b, c, d, a, x(k + 6), s34, &h4881d05 
md5_hh a, b, c, d, x(k + 9), s31, &hd9d4d039 
md5_hh d, a, b, c, x(k + 12), s32, &he6db99e5 
md5_hh c, d, a, b, x(k + 15), s33, &h1fa27cf8 
md5_hh b, c, d, a, x(k + 2), s34, &hc4ac5665 

md5_ii a, b, c, d, x(k + 0), s41, &hf4292244 
md5_ii d, a, b, c, x(k + 7), s42, &h432aff97 
md5_ii c, d, a, b, x(k + 14), s43, &hab9423a7 
md5_ii b, c, d, a, x(k + 5), s44, &hfc93a039 
md5_ii a, b, c, d, x(k + 12), s41, &h655b59c3 
md5_ii d, a, b, c, x(k + 3), s42, &h8f0ccc92 
md5_ii c, d, a, b, x(k + 10), s43, &hffeff47d 
md5_ii b, c, d, a, x(k + 1), s44, &h85845dd1 
md5_ii a, b, c, d, x(k + 8), s41, &h6fa87e4f 
md5_ii d, a, b, c, x(k + 15), s42, &hfe2ce6e0 
md5_ii c, d, a, b, x(k + 6), s43, &ha3014314 
md5_ii b, c, d, a, x(k + 13), s44, &h4e0811a1 
md5_ii a, b, c, d, x(k + 4), s41, &hf7537e82 
md5_ii d, a, b, c, x(k + 11), s42, &hbd3af235 
md5_ii c, d, a, b, x(k + 2), s43, &h2ad7d2bb 
md5_ii b, c, d, a, x(k + 9), s44, &heb86d391 

a = addunsigned(a, aa) 
b = addunsigned(b, bb) 
c = addunsigned(c, cc) 
d = addunsigned(d, dd) 
next 

md5 = lcase(wordtohex(a) & wordtohex(b) & wordtohex(c) & wordtohex(d)) 
'md5=lcase(wordtohex(b) & wordtohex(c)) 'i crop this to fit 16byte database password :d 

md5=ucase(md5) 
end function 


public function md5_16(smessage) 
m_lonbits(0) = clng(1) 
m_lonbits(1) = clng(3) 
m_lonbits(2) = clng(7) 
m_lonbits(3) = clng(15) 
m_lonbits(4) = clng(31) 
m_lonbits(5) = clng(63) 
m_lonbits(6) = clng(127) 
m_lonbits(7) = clng(255) 
m_lonbits(8) = clng(511) 
m_lonbits(9) = clng(1023) 
m_lonbits(10) = clng(2047) 
m_lonbits(11) = clng(4095) 
m_lonbits(12) = clng(8191) 
m_lonbits(13) = clng(16383) 
m_lonbits(14) = clng(32767) 
m_lonbits(15) = clng(65535) 
m_lonbits(16) = clng(131071) 
m_lonbits(17) = clng(262143) 
m_lonbits(18) = clng(524287) 
m_lonbits(19) = clng(1048575) 
m_lonbits(20) = clng(2097151) 
m_lonbits(21) = clng(4194303) 
m_lonbits(22) = clng(8388607) 
m_lonbits(23) = clng(16777215) 
m_lonbits(24) = clng(33554431) 
m_lonbits(25) = clng(67108863) 
m_lonbits(26) = clng(134217727) 
m_lonbits(27) = clng(268435455) 
m_lonbits(28) = clng(536870911) 
m_lonbits(29) = clng(1073741823) 
m_lonbits(30) = clng(2147483647) 

m_l2power(0) = clng(1) 
m_l2power(1) = clng(2) 
m_l2power(2) = clng(4) 
m_l2power(3) = clng(8) 
m_l2power(4) = clng(16) 
m_l2power(5) = clng(32) 
m_l2power(6) = clng(64) 
m_l2power(7) = clng(128) 
m_l2power(8) = clng(256) 
m_l2power(9) = clng(512) 
m_l2power(10) = clng(1024) 
m_l2power(11) = clng(2048) 
m_l2power(12) = clng(4096) 
m_l2power(13) = clng(8192) 
m_l2power(14) = clng(16384) 
m_l2power(15) = clng(32768) 
m_l2power(16) = clng(65536) 
m_l2power(17) = clng(131072) 
m_l2power(18) = clng(262144) 
m_l2power(19) = clng(524288) 
m_l2power(20) = clng(1048576) 
m_l2power(21) = clng(2097152) 
m_l2power(22) = clng(4194304) 
m_l2power(23) = clng(8388608) 
m_l2power(24) = clng(16777216) 
m_l2power(25) = clng(33554432) 
m_l2power(26) = clng(67108864) 
m_l2power(27) = clng(134217728) 
m_l2power(28) = clng(268435456) 
m_l2power(29) = clng(536870912) 
m_l2power(30) = clng(1073741824) 

dim x 
dim k 
dim aa 
dim bb 
dim cc 
dim dd 
dim a 
dim b 
dim c 
dim d 

const s11 = 7 
const s12 = 12 
const s13 = 17 
const s14 = 22 
const s21 = 5 
const s22 = 9 
const s23 = 14 
const s24 = 20 
const s31 = 4 
const s32 = 11 
const s33 = 16 
const s34 = 23 
const s41 = 6 
const s42 = 10 
const s43 = 15 
const s44 = 21 

x = converttowordarray(smessage) 

a = &h67452301 
b = &hefcdab89 
c = &h98badcfe 
d = &h10325476 

for k = 0 to ubound(x) step 16 
aa = a 
bb = b 
cc = c 
dd = d 


md5_ff a, b, c, d, x(k + 0), s11, &hd76aa478 
md5_ff d, a, b, c, x(k + 1), s12, &he8c7b756 
md5_ff c, d, a, b, x(k + 2), s13, &h242070db 
md5_ff b, c, d, a, x(k + 3), s14, &hc1bdceee 
md5_ff a, b, c, d, x(k + 4), s11, &hf57c0faf 
md5_ff d, a, b, c, x(k + 5), s12, &h4787c62a 
md5_ff c, d, a, b, x(k + 6), s13, &ha8304613 
md5_ff b, c, d, a, x(k + 7), s14, &hfd469501 
md5_ff a, b, c, d, x(k + 8), s11, &h698098d8 
md5_ff d, a, b, c, x(k + 9), s12, &h8b44f7af 
md5_ff c, d, a, b, x(k + 10), s13, &hffff5bb1 
md5_ff b, c, d, a, x(k + 11), s14, &h895cd7be 
md5_ff a, b, c, d, x(k + 12), s11, &h6b901122 
md5_ff d, a, b, c, x(k + 13), s12, &hfd987193 
md5_ff c, d, a, b, x(k + 14), s13, &ha679438e 
md5_ff b, c, d, a, x(k + 15), s14, &h49b40821 

md5_gg a, b, c, d, x(k + 1), s21, &hf61e2562 
md5_gg d, a, b, c, x(k + 6), s22, &hc040b340 
md5_gg c, d, a, b, x(k + 11), s23, &h265e5a51 
md5_gg b, c, d, a, x(k + 0), s24, &he9b6c7aa 
md5_gg a, b, c, d, x(k + 5), s21, &hd62f105d 
md5_gg d, a, b, c, x(k + 10), s22, &h2441453 
md5_gg c, d, a, b, x(k + 15), s23, &hd8a1e681 
md5_gg b, c, d, a, x(k + 4), s24, &he7d3fbc8 
md5_gg a, b, c, d, x(k + 9), s21, &h21e1cde6 
md5_gg d, a, b, c, x(k + 14), s22, &hc33707d6 
md5_gg c, d, a, b, x(k + 3), s23, &hf4d50d87 
md5_gg b, c, d, a, x(k + 8), s24, &h455a14ed 
md5_gg a, b, c, d, x(k + 13), s21, &ha9e3e905 
md5_gg d, a, b, c, x(k + 2), s22, &hfcefa3f8 
md5_gg c, d, a, b, x(k + 7), s23, &h676f02d9 
md5_gg b, c, d, a, x(k + 12), s24, &h8d2a4c8a 

md5_hh a, b, c, d, x(k + 5), s31, &hfffa3942 
md5_hh d, a, b, c, x(k + 8), s32, &h8771f681 
md5_hh c, d, a, b, x(k + 11), s33, &h6d9d6122 
md5_hh b, c, d, a, x(k + 14), s34, &hfde5380c 
md5_hh a, b, c, d, x(k + 1), s31, &ha4beea44 
md5_hh d, a, b, c, x(k + 4), s32, &h4bdecfa9 
md5_hh c, d, a, b, x(k + 7), s33, &hf6bb4b60 
md5_hh b, c, d, a, x(k + 10), s34, &hbebfbc70 
md5_hh a, b, c, d, x(k + 13), s31, &h289b7ec6 
md5_hh d, a, b, c, x(k + 0), s32, &heaa127fa 
md5_hh c, d, a, b, x(k + 3), s33, &hd4ef3085 
md5_hh b, c, d, a, x(k + 6), s34, &h4881d05 
md5_hh a, b, c, d, x(k + 9), s31, &hd9d4d039 
md5_hh d, a, b, c, x(k + 12), s32, &he6db99e5 
md5_hh c, d, a, b, x(k + 15), s33, &h1fa27cf8 
md5_hh b, c, d, a, x(k + 2), s34, &hc4ac5665 

md5_ii a, b, c, d, x(k + 0), s41, &hf4292244 
md5_ii d, a, b, c, x(k + 7), s42, &h432aff97 
md5_ii c, d, a, b, x(k + 14), s43, &hab9423a7 
md5_ii b, c, d, a, x(k + 5), s44, &hfc93a039 
md5_ii a, b, c, d, x(k + 12), s41, &h655b59c3 
md5_ii d, a, b, c, x(k + 3), s42, &h8f0ccc92 
md5_ii c, d, a, b, x(k + 10), s43, &hffeff47d 
md5_ii b, c, d, a, x(k + 1), s44, &h85845dd1 
md5_ii a, b, c, d, x(k + 8), s41, &h6fa87e4f 
md5_ii d, a, b, c, x(k + 15), s42, &hfe2ce6e0 
md5_ii c, d, a, b, x(k + 6), s43, &ha3014314 
md5_ii b, c, d, a, x(k + 13), s44, &h4e0811a1 
md5_ii a, b, c, d, x(k + 4), s41, &hf7537e82 
md5_ii d, a, b, c, x(k + 11), s42, &hbd3af235 
md5_ii c, d, a, b, x(k + 2), s43, &h2ad7d2bb 
md5_ii b, c, d, a, x(k + 9), s44, &heb86d391 

a = addunsigned(a, aa) 
b = addunsigned(b, bb) 
c = addunsigned(c, cc) 
d = addunsigned(d, dd) 
next 

'md5 = lcase(wordtohex(a) & wordtohex(b) & wordtohex(c) & wordtohex(d)) 
md5_16=lcase(wordtohex(b) & wordtohex(c)) 'i crop this to fit 16byte database password :d 

md5_16=ucase(md5_16) 
end function
%>
