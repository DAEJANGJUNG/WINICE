using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenAI;
using OpenAI.Managers;
using OpenAI.ObjectModels;
using OpenAI.ObjectModels.RequestModels;

enum Product
{
    EXIT = 0,
    OUTLOOK = 1,
    WORD = 2,
    EXCEL = 3,
    PPT = 4,
    TEST = 5
}

namespace WINICE
{

    internal static class Program
    {
        private static string ReadKeyFromConfig()
        {
            try
            {   // gpt_api_key를 읽어오는 부분
                using (StreamReader file = File.OpenText("gpt_config.json")) // gpt_config.json의 경로는 exe 파일과 동일한 경로에 존재해야 함
                {
                    using (JsonTextReader jsonReader = new JsonTextReader(file))
                    {
                        JObject json = (JObject)JToken.ReadFrom(jsonReader);
                        return json["API_KEY"].ToString();
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("Config file not found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading API key: {ex.Message}");
            }

            // 예외가 발생하거나 파일을 찾을 수 없는 경우 null 반환
            return null;
        }
        private static async Task Generate_path()
        {
            string path = Console.ReadLine();
            await RequestGPT(path);

        }

        private static async Task Init(Product item)
        {
            switch (item)
            {
                case Product.OUTLOOK:
                    // TODO : Add Parsing or Conversion API in Outlook
                    Console.WriteLine("Selected product : outlook");
                    break;
                case Product.WORD:
                    // TODO : Add Parsing or Conversion API in Word
                    Console.WriteLine("Selected product : word");
                    break;
                case Product.EXCEL:
                    // TODO : Add Parsing or Conversion API in Excel
                    Console.WriteLine("Selected product : excel");
                    break;
                case Product.PPT:
                    // TODO : Add Parsing or Conversion API in Powerpoint
                    Console.WriteLine("Selected product : powerpoint");
                    break;
                case Product.TEST:
                    break;
                default:
                    Console.WriteLine("Exit Program. Good Bye~");
                    break;
            }
            //string input_path = Get_path();
            await Generate_path();
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
            } while (selectedItem < Product.EXIT || selectedItem > Product.PPT);

            return selectedItem;
        }

        public static async Task Main(string[] args)
        {
            Product selectedItem = Menu();
            await Init(selectedItem);
        }

        public static async Task RequestGPT(string request)
        {
            var gpt3 = new OpenAIService(new OpenAiOptions()
            {
                ApiKey = ReadKeyFromConfig()
            });

            var completionResult = await gpt3.ChatCompletion.CreateCompletion(new ChatCompletionCreateRequest()
            {
                Messages = new List<ChatMessage>
                {
                    ChatMessage.FromUser(request+"와 같은 의미를 가지는 windows OS 경로를 10개 생성해줘"),
                },

                Model = Models.Gpt_3_5_Turbo,
                MaxTokens = 1000//optional
            });

            if (completionResult.Successful)
            {
                foreach (var choice in completionResult.Choices)
                {
                    Console.WriteLine("\n" + choice.Message.Content);
                }
                Console.WriteLine();
            }
            else
            {
                if (completionResult.Error == null)
                {
                    throw new Exception("Unknown Error");
                }
                Console.WriteLine($"{completionResult.Error.Code}: {completionResult.Error.Message}");
            }
        }

    }
}
