//#ifndef __stdcall
//#define __stdcall
//#endif

#ifndef NULL
#define NULL ((void*)0)
#endif

extern int __stdcall GetStdHandle(int nStdHandle);
extern int __stdcall WriteFile(int hFile,void* lpBuffer,int nNumberOfBytesToWrite ,int* lpNumberOfBytesWritten ,void* lpOverlapped );
extern int __stdcall ReadFile(int hFile,void* lpBuffer,int nNumberOfBytesToRead ,int* lpNumberOfBytesRead ,void* lpOverlapped );

int __stdcall Factorial(int n){
int i=0,j=0;
j=1;
//for(i=1;i<=n;i++) j*=i;
return j;
}

void __stdcall PrintInteger(int n){
int h=0,i=0;
int temp1,temp2;
h = GetStdHandle(-12);
if(n<0){
n=-n;
temp1=45;
temp2=0;
WriteFile(h,&temp1,1,&temp2,NULL);
}
if(n==0){
temp1=0x30;
temp2=0;
WriteFile(h,&temp1,1,&temp2,NULL);
}else{
i=n/10;
if(i>0) PrintInteger(i);
temp1=0x30|(n%10);
temp2=0;
WriteFile(h,&temp1,1,&temp2,NULL);
}
}

int main(){
int h=0,i=0,j=0,c=0;
h=GetStdHandle(-10);
ReadFile(h,&c,1,&j,NULL);
PrintInteger(Factorial(c&0xF));
return 0;
}
