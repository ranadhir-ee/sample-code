//Pre-processor directive
#include <stdio.h>
// Main function
int main()
{
	//Declare variables
	int flyPack;
	int wifiPack;

	char pickService;

	int carPack = 0;
	int hotelPack;
	
	int packCount; 
	float total;
	
	//Print welcome screen
	printf("Hello and welcome to the AMCD Travels  and  Tour service : \n \n");
	
	//Prompt to user for flights subscription
	printf("Would you like to subscribe to the flights service? Type 'Y' for yes or 'N' for no : \n ");
	scanf(" %c", &pickService);
	
	//Prompt to user for flight package type
	if(pickService == 'Y' || pickService == 'y')
	{
		packCount++;
		
		printf("Please pick your flight package by typing the package number. Type '1', '2' or '3' to choose : \n ");
		scanf("%d", &flyPack);
		
		//Prompt to user for optional wifi subscription
		printf("Next, please choose a wifi pakage. Type '1' for first 3 flights or '2' for all flights : \n ");
		scanf("%d", &wifiPack);
	}
	
	//Prompt to user for rental cars subscription
	printf("\n \nWould you like to subscribe to the rental cars service? Type 'Y' for yes or 'N' for no : \n ");
	scanf(" %c", &pickService);
	
	//Prompt to user for rental car package type
	if(pickService == 'Y' || pickService == 'y')
	{
		packCount++;
		
		printf("Please pick your rental car package by typing the package number. Type '1', '2' or '3' to choose : \n ");
		scanf("%d", &carPack);
	}
	
	
	//Prompt to user for flight Hotel subscription 
	printf("\n\nFinally, would you like to subscribe to the hotel service? Type 'Y' for yes or 'N' for no \n : ");
	scanf(" %c", &pickService);
	
	//Prompt to user for flight Hotel package type
	if(pickService == 'Y' || pickService == 'y')
	{
		packCount++;
		
		printf("Which Package would you like to choose? Type '1' or '2' \n ");
		scanf("%d", &hotelPack);
	}
	
	
	//Print Receipt
	//Print header
	printf("\n\n---------------------------------------------------------------------------- \nAMCD TOURS AND TRAVELSINC Bill Details \n---------------------------------------------------------------------------- \n");
	
	printf("Service                              Package                 Cost \n\n");
	
	//Receipt for Flights
	if(flyPack)
	{	//print Flights
		printf("Flights                                 %d                    ", flyPack);
		switch(flyPack)
		{	//print cost
			case(1):
				printf("$%6.2lf \n", 29.99);
				total += 29.99;
				break;
			//print cost
			case(2):
				printf("$%6.2lf \n", 39.99);
				total += 39.99;
				break;
			//print cost
			case(3):
				printf("$%6.2lf \n", 49.99);
				total += 49.99;
				break;
		}
		//print WiFi
		printf("\t First 3 Flights:               %d                    ", wifiPack);
		switch(wifiPack)
		{	//print cost
			case(1):
				printf("$%6.2lf \n", 2.99);
				total += 2.99;
				break;
			//print cost
			case(2):
				printf("$%6.2lf \n", 6.99);
				total += 6.99;
				break;
		}
	}

	//Reciept for Rental Cars
	if(carPack)
	{	//print Rental Car
		printf("Rental Car                              %d                    ", carPack);
		switch(carPack)
		{	//print cost
			case(1):
				printf("$%6.2lf \n", 19.99);
				total += 19.99;
				break;
			//print cost
			case(2):
				printf("$%6.2lf \n", 29.99);
				total += 29.99;
				break;
			//print cost
			case(3):
				printf("$%6.2lf \n", 39.99);
				total += 39.99;
				break;
		}
	}

	//Receipt for hotels
	if(hotelPack)
	{	//print Hotel
		printf("Hotel                                   %d                    ", hotelPack);
		switch(hotelPack)
		{	//print cost
			case(1):
				printf("$%6.2lf \n", 59.99);
				total += 59.99;
				break;
			//print cost
			case(2):
				printf("$%6.2lf \n", 99.99);
				total += 99.99;
				break;
		}
	}
	
	printf("Subtotal                                                     $%6.2lf\n", total); //subtotal
	
	if(packCount > 1)//if more than 1 package selected
	{	//Bundle discount
		printf("Bundle Discount                         %d%%                   $%6.2lf\n", 5, 0.05*total); 
		total -= 0.05*total;
	}
	
	printf("Total Before Tax                                             $%6.2lf\n", total); //before tax cost
	printf("Tax                                     %d%%                   $%6.2lf\n", 8, 0.08*total); // tax amount
	total += 0.08*total;
	//after tax cost
	printf("Amount Due                                                   $%6.2lf \n ----------------------------------------------------------------------------", total);	//final total

}
