#include <stdio.h>
#include <Windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <conio.h>


int main(){

	int age, gugan;

	SetConsoleTitle(TEXT("���-����ö��ǥ����ȭü��-����ũ�μ���Ʈ�ü������������-����10.0.10166W1nd0ws"));

	printf("������ �Է��Ͻö��{1,2����}");

	scanf("%d", &gugan);

	switch(gugan){

		case 1:
			printf("���̹����� �Է��Ͻö��{������:1}{û�ҳ�:2}{����:3}{65�� �̻�:4} : ");

	         scanf( "%d", &age);

			switch(age){

				case 1:
					printf("���� ������ʿ� �̶��.");
					break;
				case 2:
					printf("���� õ���ʿ� �̶��.");
					break;
				case 3:
					printf("���� õ���� �̶��.");
					break;
				case 4:
					printf("���� ������ʿ� �̶��.");
					break;
				default:
				 	printf("�׵� ���̴� ������.");
				}
			break;
		case 2:
			printf("���̹����� �Է��Ͻö��{������:1}{û�ҳ�:2}{����:3}{65�� �̻�:4} :");

			scanf( "%d", &age);

			switch(age){

				case 1:
					printf("���� ĥ����ʿ� �̶��.");
					break;
				case 2:
					printf("���� õ����ʿ� �̶��.");
					break;
				case 3:
					printf("���� õ���� �̶��.");
					break;
				case 4:
					printf("���� ĥ����ʿ� �̶��.");
					break;

				default:
				 	printf("�׵� ���̴� ������.");
				}
			break;
		default:
			printf("�׵� ������ �ƿ� ������ ������.\n");
			printf("�Ӹ��� ������ �ȹٷ� �����غ��ö��.");
			return 0;
		}





	}





