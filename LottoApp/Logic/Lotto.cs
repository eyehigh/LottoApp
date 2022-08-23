using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LottoApp.Logic
{
    internal class Lotto
    {
        void TestBitArray()
        {
            // 8개의 비트를 갖는 BitArray 객체 생성
            BitArray ba1 = new BitArray(8);

            // ## 비트 쓰기 ##
            ba1.Set(0, true);
            ba1.Set(1, true);

            PrintBits(ba1);  // 11000000

            BitArray ba2 = new BitArray(8);
            ba2[1] = true;
            ba2[2] = true;
            ba2[3] = true;
            PrintBits(ba2);  // 01110000

            // ## 비트 읽기 ##
            bool b1 = ba1.Get(0); // true
            bool b2 = ba2[4];     // false


            // ## BitArray 비트 연산 ## 
            // OR (ba1 | ba2) 결과를 ba1 에 
            ba1.Or(ba2);
            PrintBits(ba1);  // 11110000

            // AND (ba1 & ba2) 결과를 ba1 에 
            ba1.And(ba2);
            PrintBits(ba1);  // 01110000

            ba1.Xor(ba2);
            ba1.Not();


            // ## 기타 BitArray 생성 방법 ##
            // bool[] 로 생성
            var bools = new bool[] { true, true, false, false };
            BitArray ba3 = new BitArray(bools);

            // byte[] 로 생성
            var bytes = new byte[] { 0xFF, 0x11 };
            BitArray ba4 = new BitArray(bytes);
        }
        void PrintBits(BitArray ba)
        {
            for (int i = 0; i < ba.Count; i++)
            {
                Console.Write(ba[i] ? "1" : "0");
            }
            Console.WriteLine();
        }
        public BitArray ToBitArray(int[] numbers)
        {
            BitArray ba = new BitArray(45);
            ba[ numbers[0] - 1 ] = true;
            ba[ numbers[1] - 1 ] = true;
            ba[ numbers[2] - 1 ] = true;
            ba[ numbers[3] - 1 ] = true;
            ba[ numbers[4] - 1 ] = true;
            ba[ numbers[5] - 1 ] = true;
            return ba;
        }
        public string BitToString(BitArray ba)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < ba.Count; i++)
            {
                sb.Append(ba[i] ? "1" : "0");
            }
            return sb.ToString();
        }
        public BitArray StringToBit(string str)
        {
            BitArray ba = new BitArray(45);
            for(int i =0; i<45; i++)
            {
                string s = str.Substring(i, 1);
                if (s.Equals("1"))
                {
                    ba[i] = true;
                }
                else
                {
                    ba[i] = false;
                }
            }
            return ba;
        }
        public void CheckLogic1()
        {

        }
    }
}
