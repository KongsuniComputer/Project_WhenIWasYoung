#include <stdio.h>
int main()
{
	int n, i, j, sum;
	int cnt = 0;

	printf("���� �Է��ϼ���\n");
	scanf_s("%d", &n);

	for (i = 1; i <= n; i++)
	{
		sum = 0;
		for (j = 1; j<i; j++)
		{
			if (i % j == 0) sum += j;
		}
		if (sum == i)
		{
			printf("%d ", i); 
			cnt++; 
		}
	}
	printf("\n�������� %d�� �Դϴ�.\n", cnt);

	return 0;
}

// Visual Studio Express 2013 ���� �ۼ���