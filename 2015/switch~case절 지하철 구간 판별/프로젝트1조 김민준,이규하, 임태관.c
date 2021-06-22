#include <stdio.h>
#include <Windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <conio.h>


int main(){

	int age, gugan;

	SetConsoleTitle(TEXT("평양-지하철매표정보화체계-마이크로소프트운영체제엠에스도스-버전10.0.10166W1nd0ws"));

	printf("구간을 입력하시라우{1,2구간}");

	scanf("%d", &gugan);

	switch(gugan){

		case 1:
			printf("나이범위를 입력하시라우{려린이:1}{청소년:2}{성인:3}{65세 이상:4} : ");

	         scanf( "%d", &age);

			switch(age){

				case 1:
					printf("돈은 륙백오십원 이라우.");
					break;
				case 2:
					printf("돈은 천오십원 이라우.");
					break;
				case 3:
					printf("돈은 천삼백원 이라우.");
					break;
				case 4:
					printf("돈은 륙백오십원 이라우.");
					break;
				default:
				 	printf("그딴 나이는 없슴메.");
				}
			break;
		case 2:
			printf("나이범위를 입력하시라우{려린이:1}{청소년:2}{성인:3}{65세 이상:4} :");

			scanf( "%d", &age);

			switch(age){

				case 1:
					printf("돈은 칠백오십원 이라우.");
					break;
				case 2:
					printf("돈은 천백오십원 이라우.");
					break;
				case 3:
					printf("돈은 천사백원 이라우.");
					break;
				case 4:
					printf("돈은 칠백오십원 이라우.");
					break;

				default:
				 	printf("그딴 나이는 없슴메.");
				}
			break;
		default:
			printf("그딴 구간은 아예 있지도 않으메.\n");
			printf("머리가 있으면 똑바로 생각해보시라우.");
			return 0;
		}





	}





