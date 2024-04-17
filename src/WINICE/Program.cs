using System;
using System.ComponentModel.Design;
using System.Xml;

enum Product
{
    EXIT = 0,
    OUTLOOK = 1,
    WORD = 2,
    EXCEL = 3,
    PPT = 4
} 

class Program
{
    private static void Init(Product item)
    {
        switch (item)
        {
            case Product.OUTLOOK:
                break;
            case Product.WORD:
                break;
            case Product.EXCEL:
                break;
            case Product.PPT:
                break;
            default:
                break;
        }
    }
    public static Product Menu()
    {
        Product selectedItem;
        do
        {
            Console.WriteLine(@"
          ___                       ___                       ___           ___     
         /\  \                     /\  \                     /\__\         /\__\    
        _\:\  \       ___          \:\  \       ___         /:/  /        /:/ _/_   
       /\ \:\  \     /\__\          \:\  \     /\__\       /:/  /        /:/ /\__\  
      _\:\ \:\  \   /:/__/      _____\:\  \   /:/__/      /:/  /  ___   /:/ /:/ _/_ 
     /\ \:\ \:\__\ /::\  \     /::::::::\__\ /::\  \     /:/__/  /\__\ /:/_/:/ /\__\
     \:\ \:\/:/  / \/\:\  \__  \:\~~\~~\/__/ \/\:\  \__  \:\  \ /:/  / \:\/:/ /:/  /
      \:\ \::/  /   ~~\:\/\__\  \:\  \        ~~\:\/\__\  \:\  /:/  /   \::/_/:/  / 
       \:\/:/  /       \::/  /   \:\  \          \::/  /   \:\/:/  /     \:\/:/  /  
        \::/  /        /:/  /     \:\__\         /:/  /     \::/  /       \::/  /   
         \/__/         \/__/       \/__/         \/__/       \/__/         \/__/    
      ");

            Console.WriteLine("Please Select The Object to Analysis");
            Console.WriteLine("== Object List ==");
            Console.WriteLine(@"
      [0] Exit
      [1] Outlook
      [2] Word
      [3] Excel
      [4] Power Point
      ");

            Console.Write("Select item : ");
            if (int.TryParse(Console.ReadLine(), out int item))
            {
                selectedItem = (Product)item;
            }
            else
            {
                Console.WriteLine("Invalid input");
                selectedItem = Product.EXIT;
            }
        } while (selectedItem < Product.EXIT ||  selectedItem > Product.PPT);

        return selectedItem;
    }
    public static void Main(string[] args)
    {
        Product selectedItem = Menu();
        Init(selectedItem);
    }

}
