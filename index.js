import chalk from 'chalk';
import { Wallet, ethers } from 'ethers';
import moment from 'moment';
import readlineSync from 'readline-sync';
import ExcelJS from 'exceljs'; // Import ExcelJS

// Function to create a new Ethereum account
function createAccountETH() {
  const wallet = ethers.Wallet.createRandom();
  const privateKey = wallet.privateKey;
  const publicKey = wallet.publicKey;
  const mnemonicKey = wallet.mnemonic.phrase;

  return { privateKey, publicKey, mnemonicKey };
}

// Main function using async IIFE (Immediately Invoked Function Expression)
(async () => {
  try {
    // Get the total number of wallets to create from user input
    const totalWallet = readlineSync.question(
      chalk.yellow('Input how many wallets you want to create: ')
    );

    let count = 1;

    // If the user entered a valid number greater than 1, set the count
    if (totalWallet > 1) {
      count = totalWallet;
    }

    // Initialize a new Excel workbook and sheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Wallets');

    // Add headers to the worksheet
    worksheet.columns = [
      { header: 'Address', key: 'address', width: 42 },
      { header: 'Private Key', key: 'privateKey', width: 66 },
      { header: 'Mnemonic', key: 'mnemonic', width: 60 }
    ];

    // Create the specified number of wallets
    while (count > 0) {
      const createWalletResult = createAccountETH();
      const theWallet = new Wallet(createWalletResult.privateKey);

      if (theWallet) {
        // Add wallet details to the Excel sheet
        worksheet.addRow({
          address: theWallet.address,
          privateKey: createWalletResult.privateKey,
          mnemonic: createWalletResult.mnemonicKey
        });

        // Display success message with the wallet address and timestamp
        console.log(
          chalk.green(
            `[${moment().format('HH:mm:ss')}] => ` +
              'Wallet created...! Your address: ' + theWallet.address
          )
        );
      }

      count--;
    }

    // Write the Excel file
    await workbook.xlsx.writeFile('result.xlsx');

    // Display final message after creating all wallets
    setTimeout(() => {
      console.log(
        chalk.green(
          'All wallets have been created. Check result.xlsx to see the address, mnemonic, and private key.'
        )
      );
    }, 3000);

    return;
  } catch (error) {
    // Display error message if an error occurs
    console.log(chalk.red('Your program encountered an error! Message: ' + error));
  }
})();
