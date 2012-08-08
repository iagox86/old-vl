#include "stdafx.h"
#include <string.h>
#include <stdio.h>

FILE *SafeOpen(char *name, char *mode);
void SafeClose(FILE *file);
void ProcessTheFile(FILE *InputFile, FILE *OutputFile);

int main(int argc, char* argv[])
{
	//Assumes that argv[1] is the name of the input file
	//Assumes that argv[2] is the name of the output file
	//Error is returned if there are no arguments

	//Files:
	FILE *InputFile;
	FILE *OutputFile;

	if(argc != 3)
	{
		fprintf(stderr, "Error: This program requires 2 parameters\nvL2.exe (inputfile) (outputfile)", argc);
		return 0xBAD; // My favorite error message :D
	}

	InputFile = SafeOpen(argv[1], "rb");
	OutputFile = SafeOpen(argv[2], "w");

	if((InputFile && OutputFile) == NULL)
	{
		fprintf(stderr, "ERROR: One or both of the files were unreadable");
		return 0xBAD; //My favorite error :D
	}

	ProcessTheFile(InputFile, OutputFile);

	SafeClose(InputFile);
	SafeClose(OutputFile);

	return 0;
}


void ProcessTheFile(FILE *InputFile, FILE *OutputFile)
{
	bool StillGoing = true;
	bool Good;
	bool Changed;

	short AtoZ;
		//in AtoZ:
		//0 = unknown
		//1 = A --> Z
		//2 = Z --> A

	char *ReturnString = "";
	
	char Letter;
	char PreviousLetter;

	char word[1024];
	
	int position = 0;

//	while (StillGoing)
	while(StillGoing)
	{
		Good = true;
		PreviousLetter = 0;
		AtoZ = 0;

		position = 0;
		word[0] = 0;

		do
		{
			//fscanf(InputFile, "%c\n", &Letter);
			Letter = getc(InputFile);
			Changed = false;

			if(Letter == 13) 
			{
				fscanf(InputFile, "\n");
				Letter = ' ';
			}
			else if(Letter == -1)
			{
				StillGoing = false;
				Letter = ' ';
			}
			else if(Letter >= 'a' && Letter <= 'z')
			{
				Letter = Letter - 0x20;
				Changed = true;
			}

			if(Letter >= 'A' && Letter <= 'Z')
			{
				//printf("%c -- %d\n", Letter, Letter);
				if (PreviousLetter)
				{
					if(!(AtoZ))
					{
						if(Letter>PreviousLetter)
							AtoZ = 1;
						else if(Letter<PreviousLetter)
							AtoZ = 2;
					}
					else
					{
						if((Letter>PreviousLetter && AtoZ != 1) || (Letter < PreviousLetter && AtoZ != 2))
							Good=false;
					}
				}
				PreviousLetter = Letter;
			}
			
			if(Changed)
				Letter += 0x20;

			word[position] = Letter;
			word[++position] = 0;
			//printf("%s\n", word);
		}
		while(Letter != ' ');

		if(Good)
		{
			fprintf(OutputFile, "%s\n", word);
			fprintf(stdout, "%s\n", word);
		}
	}
}





//I programmed this pair of functions for a school project:
FILE *SafeOpen(char *name, char *mode)
{
   FILE *file;

   if((file = fopen(name, mode)) == NULL)
   {
      fprintf(stderr, "ERROR: File %s could not be opened.\n\n", name);
      return NULL;
   }

   return file;
}

void SafeClose(FILE *file)
{
   if (fclose(file) == EOF)
   {
      fprintf(stderr, "ERROR: File cannot be closed\n(error ignored)");
   }
}

